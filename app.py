#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP 与 CBS 对账 Web 服务：提供网页上传 Excel、执行对账并下载结果。
运行后本机及局域网内可通过 http://本机IP:5001 访问。
"""

import base64
import io
import socket
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request

from backend_reconcile_sap_cbs import (
    load_cbs_detail_from_bytes,
    load_sap_detail_from_bytes,
    reconcile,
    report_as_string,
)

app = Flask(__name__)

# 内联 HTML 模板，避免单独 templates 目录依赖
INDEX_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SAP 与 CBS 对账</title>
  <style>
    body { font-family: "Microsoft YaHei", sans-serif; max-width: 640px; margin: 2rem auto; padding: 0 1rem; }
    h1 { font-size: 1.4rem; }
    .field { margin: 1rem 0; }
    .field label { display: block; margin-bottom: 0.3rem; font-weight: 500; }
    input[type="file"] { width: 100%; }
    button { padding: 0.5rem 1.2rem; background: #2563eb; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 1rem; }
    button:disabled { opacity: 0.6; cursor: not-allowed; }
    #status { margin: 1rem 0; min-height: 1.2rem; color: #666; }
    #report { white-space: pre-wrap; background: #f5f5f5; padding: 1rem; border-radius: 6px; margin: 1rem 0; font-size: 0.9rem; display: none; }
    #error { color: #b91c1c; margin: 1rem 0; display: none; }
    .downloads { margin-top: 1rem; display: none; }
    .downloads button { margin-right: 0.5rem; margin-bottom: 0.5rem; background: #059669; }
    .downloads button.secondary { background: #6b7280; }
  </style>
</head>
<body>
  <h1>SAP 与 CBS 对账</h1>
  <p>上传 CBS 明细与 SAP 明细 Excel，点击对账后可在页面查看报告并下载结果。</p>
  <form id="form">
    <div class="field">
      <label for="cbs">CBS 明细（.xlsx / .xls）</label>
      <input type="file" id="cbs" name="cbs_file" accept=".xlsx,.xls" required>
    </div>
    <div class="field">
      <label for="sap">SAP 明细（.xlsx / .xls）</label>
      <input type="file" id="sap" name="sap_file" accept=".xlsx,.xls" required>
    </div>
    <button type="submit" id="btn">开始对账</button>
  </form>
  <div id="status"></div>
  <div id="error"></div>
  <pre id="report"></pre>
  <div class="downloads" id="downloads">
    <button type="button" id="dlExcel" class="secondary">下载对账结果.xlsx</button>
    <button type="button" id="dlReport">下载对账报告.txt</button>
  </div>
  <script>
    const form = document.getElementById("form");
    const status = document.getElementById("status");
    const error = document.getElementById("error");
    const report = document.getElementById("report");
    const downloads = document.getElementById("downloads");
    const btn = document.getElementById("btn");
    const dlExcel = document.getElementById("dlExcel");
    const dlReport = document.getElementById("dlReport");

    let lastResult = null;

    function showError(msg) {
      error.textContent = msg;
      error.style.display = "block";
      report.style.display = "none";
      downloads.style.display = "none";
    }
    function hideError() {
      error.style.display = "none";
    }

    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      const cbsFile = document.getElementById("cbs").files[0];
      const sapFile = document.getElementById("sap").files[0];
      if (!cbsFile || !sapFile) {
        showError("请同时选择 CBS 明细和 SAP 明细文件。");
        return;
      }
      hideError();
      status.textContent = "对账中…";
      btn.disabled = true;
      downloads.style.display = "none";
      report.style.display = "none";

      const fd = new FormData();
      fd.append("cbs_file", cbsFile);
      fd.append("sap_file", sapFile);

      try {
        const res = await fetch("/reconcile", { method: "POST", body: fd });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) {
          showError(data.error || data.message || "请求失败 " + res.status);
          status.textContent = "";
          btn.disabled = false;
          return;
        }
        lastResult = data;
        report.textContent = data.report || "";
        report.style.display = "block";
        downloads.style.display = "block";
        status.textContent = "对账完成，可下载结果。";
      } catch (err) {
        showError("网络错误：" + err.message);
        status.textContent = "";
      }
      btn.disabled = false;
    });

    function b64ToBlob(b64, mime) {
      const bin = atob(b64);
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      return new Blob([arr], { type: mime });
    }
    function downloadBlob(blob, filename) {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
      URL.revokeObjectURL(a.href);
    }

    dlExcel.addEventListener("click", () => {
      if (!lastResult || !lastResult.excel_base64) return;
      const blob = b64ToBlob(lastResult.excel_base64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      downloadBlob(blob, "CBS明细_对账结果.xlsx");
    });
    dlReport.addEventListener("click", () => {
      if (!lastResult || !lastResult.report) return;
      const blob = new Blob([lastResult.report], { type: "text/plain;charset=utf-8" });
      downloadBlob(blob, "对账报告.txt");
    });
  </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(INDEX_HTML)


@app.route("/reconcile", methods=["POST"])
def do_reconcile():
    cbs_file = request.files.get("cbs_file")
    sap_file = request.files.get("sap_file")
    if not cbs_file or not sap_file:
        return jsonify({"error": "请同时上传 CBS 明细和 SAP 明细文件。"}), 400

    try:
        cbs_bytes = cbs_file.read()
        sap_bytes = sap_file.read()
    except Exception as e:
        return jsonify({"error": f"读取上传文件失败：{e}"}), 400

    try:
        cbs_df = load_cbs_detail_from_bytes(cbs_bytes)
        sap_df = load_sap_detail_from_bytes(sap_bytes)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    try:
        result = reconcile(cbs_df, sap_df)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"对账过程出错：{e}"}), 500

    report_text = report_as_string(result)

    buffer = io.BytesIO()
    result.to_excel(buffer, index=False)
    buffer.seek(0)
    excel_b64 = base64.b64encode(buffer.getvalue()).decode("ascii")

    return jsonify({"report": report_text, "excel_base64": excel_b64})


if __name__ == "__main__":
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        lan_ip = s.getsockname()[0]
        s.close()
    except Exception:
        lan_ip = "本机IP"
    print("局域网：http://{}:5001".format(lan_ip))
    app.run(host="0.0.0.0", port=5001)
