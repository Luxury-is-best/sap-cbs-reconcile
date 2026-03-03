#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
后端对账核心模块（上线文件名）。

为保持与现有逻辑完全一致，这里直接复用 `reconcile_sap_cbs.py` 的实现与命令行入口。
"""

from reconcile_sap_cbs import (  # noqa: F401
    PROJECT_DIR,
    fuzzy_find_cbs_file,
    fuzzy_find_sap_file,
    load_cbs_detail,
    load_sap_detail,
    load_cbs_detail_from_bytes,
    load_sap_detail_from_bytes,
    reconcile,
    report_as_string,
    write_report,
    main,
)


if __name__ == "__main__":
    main()
