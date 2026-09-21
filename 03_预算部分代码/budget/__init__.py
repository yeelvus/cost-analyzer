# -*- coding: utf-8 -*-
"""
预算处理包（广联达人材机 + 看板 + 责任成本）。

用法：
  python -m budget              # 交互菜单 / 默认全流程
  python -m budget glodon       # 仅广联达 raw→std+总汇总
  python -m budget dashboard    # 仅生成看板
  python -m budget all          # 广联达 + 看板
  python -m budget responsibility  # 责任成本

扩展新功能：在 budget/ 下新增模块，并在 budget/__main__.py 注册子命令即可。
"""

__version__ = "1.0.0"
