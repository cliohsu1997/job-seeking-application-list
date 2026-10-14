"""
Catalog universities + economics units for:
  Guangdong (excl. Shenzhen), Fujian, Zhejiang, Shenzhen

Each region is an independent folder under data/economics/:
  guangdong/  fujian/  zhejiang/  shenzhen/

Fetches department/career pages, writes Application-column listings + clean xlsx.
"""
from __future__ import annotations

import argparse
import gzip
import json
import re
import ssl
import time
import zlib
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CTX = ssl._create_unverified_context()
TODAY = date.today().isoformat()
SLEEP = 0.35

REGIONS = ("guangdong", "fujian", "zhejiang", "shenzhen")

# Primary key = Chinese name (higher hit-rate for CN sites). English kept for Application columns.
# Shenzhen schools ONLY under shenzhen/ (not duplicated in guangdong/).
UNIVERSITIES: dict[str, list[dict]] = {
    "guangdong": [
        {
            "name_zh": "中山大学",
            "name_en": "Sun Yat-sen University",
            "city": "广州",
            "home": "https://www.sysu.edu.cn/",
            "career": "https://rcb.sysu.edu.cn/zp",
            "postdoc": "https://lingnan.sysu.edu.cn/article/2580",
            "econ_units": [
                {
                    "unit_zh": "岭南学院",
                    "unit_en": "Lingnan College",
                    "url": "https://lingnan.sysu.edu.cn/",
                    "career": "https://lingnan.sysu.edu.cn/Recruitment",
                }
            ],
        },
        {
            "name_zh": "暨南大学",
            "name_en": "Jinan University",
            "city": "广州",
            "home": "https://www.jnu.edu.cn/",
            "career": "https://hrdam.jnu.edu.cn/264/list.htm",
            "postdoc": "https://hrdam.jnu.edu.cn/2026/0413/c264a853500/page.htm",
            "econ_units": [
                {
                    "unit_zh": "经济与社会研究院",
                    "unit_en": "IESR",
                    "url": "https://iesr.jnu.edu.cn/",
                    "career": "https://iesr.jnu.edu.cn/",
                },
                {
                    "unit_zh": "经济学院",
                    "unit_en": "College of Economics",
                    "url": "https://ec.jnu.edu.cn/",
                },
            ],
        },
        {
            "name_zh": "华南理工大学",
            "name_en": "South China University of Technology",
            "city": "广州",
            "home": "https://www2.scut.edu.cn/",
            "career": "https://www2.scut.edu.cn/hr/",
            "postdoc": "https://www2.scut.edu.cn/hr/",
            "econ_units": [
                {
                    "unit_zh": "经济与金融学院",
                    "unit_en": "School of Economics and Finance",
                    "url": "https://www2.scut.edu.cn/economy/",
                }
            ],
        },
        {
            "name_zh": "华南师范大学",
            "name_en": "South China Normal University",
            "city": "广州",
            "home": "https://www.scnu.edu.cn/",
            "career": "https://rsc.scnu.edu.cn/",
            "postdoc": "https://rsc.scnu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济与管理学院",
                    "unit_en": "School of Economics & Management",
                    "url": "https://em.scnu.edu.cn/",
                    "career": "https://rsc.scnu.edu.cn/rencaizhaopin/",
                }
            ],
        },
        {
            "name_zh": "广东外语外贸大学",
            "name_en": "Guangdong University of Foreign Studies",
            "city": "广州",
            "home": "https://www.gdufs.edu.cn/",
            "career": "https://rsc.gdufs.edu.cn/",
            "postdoc": "https://rsc.gdufs.edu.cn/info/1017/12559.htm",
            "econ_units": [
                {
                    "unit_zh": "经济贸易学院",
                    "unit_en": "School of Economics and Trade",
                    "url": "https://seib.gdufs.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "广东财经大学",
            "name_en": "Guangdong University of Finance and Economics",
            "city": "广州",
            "home": "https://www.gdufe.edu.cn/",
            "career": "https://rsc.gdufe.edu.cn/",
            "postdoc": "https://rsc.gdufe.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.gdufe.edu.cn/",
                    "career": "http://rsc.gdufe.edu.cn/bshzp/list.htm",
                }
            ],
        },
        {
            "name_zh": "广州大学",
            "name_en": "Guangzhou University",
            "city": "广州",
            "home": "https://www.gzhu.edu.cn/",
            "career": "https://rsc.gzhu.edu.cn/",
            "postdoc": "https://rsc.gzhu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济与统计学院",
                    "unit_en": "School of Economics and Statistics",
                    "url": "https://jjxy.gzhu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "广东工业大学",
            "name_en": "Guangdong University of Technology",
            "city": "广州",
            "home": "https://www.gdut.edu.cn/",
            "career": "https://hr.gdut.edu.cn/",
            "postdoc": "https://hr.gdut.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.gdut.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "华南农业大学",
            "name_en": "South China Agricultural University",
            "city": "广州",
            "home": "https://www.scau.edu.cn/",
            "career": "https://hr.scau.edu.cn/",
            "postdoc": "https://hr.scau.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济管理学院",
                    "unit_en": "College of Economics and Management",
                    "url": "https://cem.scau.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "广东金融学院",
            "name_en": "Guangdong University of Finance",
            "city": "广州",
            "home": "https://rsc.gduf.edu.cn/",
            "career": "https://rsc.gduf.edu.cn/",
            "postdoc": "https://rsc.gduf.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济贸易学院",
                    "unit_en": "School of Economics and Trade",
                    "url": "https://rsc.gduf.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "广东海洋大学",
            "name_en": "Guangdong Ocean University",
            "city": "湛江",
            "home": "https://www.gdou.edu.cn/",
            "career": "https://rsc.gdou.edu.cn/",
            "postdoc": "https://rsc.gdou.edu.cn/",
            "econ_units": [],
        },
        {
            "name_zh": "北京师范大学香港浸会大学联合国际学院",
            "name_en": "BNU-HKBU UIC",
            "city": "珠海",
            "home": "https://uic.edu.cn/",
            "career": "https://uic.edu.cn/",
            "postdoc": "https://uic.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "工商管理学院",
                    "unit_en": "Faculty of Business and Management",
                    "url": "https://uic.edu.cn/",
                }
            ],
        },
    ],
    "fujian": [
        {
            "name_zh": "厦门大学",
            "name_en": "Xiamen University",
            "city": "厦门",
            "home": "https://www.xmu.edu.cn/",
            "career": "https://jobs.xmu.edu.cn/",
            "postdoc": "https://jobs.xmu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://soe.xmu.edu.cn/",
                },
                {
                    "unit_zh": "经济学系",
                    "unit_en": "Department of Economics",
                    "url": "https://econ.xmu.edu.cn/",
                },
                {
                    "unit_zh": "王亚南经济研究院",
                    "unit_en": "WISE",
                    "url": "https://wise.xmu.edu.cn/",
                },
            ],
        },
        {
            "name_zh": "福州大学",
            "name_en": "Fuzhou University",
            "city": "福州",
            "home": "https://www.fzu.edu.cn/",
            "career": "https://rsc.fzu.edu.cn/xiaowai/zxzp.htm",
            "postdoc": "https://rsc.fzu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济与管理学院",
                    "unit_en": "School of Economics and Management",
                    "url": "https://jgxy.fzu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "福建师范大学",
            "name_en": "Fujian Normal University",
            "city": "福州",
            "home": "https://www.fjnu.edu.cn/",
            "career": "https://rsc.fjnu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.fjnu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "华侨大学",
            "name_en": "Huaqiao University",
            "city": "泉州/厦门",
            "home": "https://www.hqu.edu.cn/",
            "career": "https://rsc.hqu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济与金融学院",
                    "unit_en": "School of Economics and Finance",
                    "url": "https://ef.hqu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "集美大学",
            "name_en": "Jimei University",
            "city": "厦门",
            "home": "https://www.jmu.edu.cn/",
            "career": "https://rsc.jmu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "财经学院",
                    "unit_en": "School of Finance and Economics",
                    "url": "https://finance.jmu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "福建农林大学",
            "name_en": "Fujian Agriculture and Forestry University",
            "city": "福州",
            "home": "https://www.fafu.edu.cn/",
            "career": "https://rsc.fafu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济与管理学院",
                    "unit_en": "School of Economics and Management",
                    "url": "https://jjgl.fafu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "闽南师范大学",
            "name_en": "Minnan Normal University",
            "city": "漳州",
            "home": "https://www.mnnu.edu.cn/",
            "career": "https://rsc.mnnu.edu.cn/",
            "econ_units": [],
        },
        {
            "name_zh": "福建江夏学院",
            "name_en": "Fujian Jiangxia University",
            "city": "福州",
            "home": "https://www.fjjxu.edu.cn/",
            "career": "https://rsc.fjjxu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济贸易学院",
                    "unit_en": "School of Economics and Trade",
                    "url": "https://jmxy.fjjxu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "泉州师范学院",
            "name_en": "Quanzhou Normal University",
            "city": "泉州",
            "home": "https://www.qztc.edu.cn/",
            "career": "https://rsc.qztc.edu.cn/",
            "econ_units": [],
        },
        {
            "name_zh": "仰恩大学",
            "name_en": "Yang-En University",
            "city": "泉州",
            "home": "https://www.yeu.edu.cn/",
            "career": "https://www.yeu.edu.cn/",
            "econ_units": [],
        },
    ],
    "zhejiang": [
        {
            "name_zh": "浙江大学",
            "name_en": "Zhejiang University",
            "city": "杭州",
            "home": "https://www.zju.edu.cn/",
            "career": "http://talent.zju.edu.cn/",
            "postdoc": "http://talent.zju.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "http://www.cec.zju.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "浙江财经大学",
            "name_en": "Zhejiang University of Finance and Economics",
            "city": "杭州",
            "home": "https://www.zufe.edu.cn/",
            "career": "https://rsc.zufe.edu.cn/info/1073/3604.htm",
            "postdoc": "https://rsc.zufe.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.zufe.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "浙江工商大学",
            "name_en": "Zhejiang Gongshang University",
            "city": "杭州",
            "home": "https://www.zjgsu.edu.cn/",
            "career": "https://rsc.zjgsu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.zjgsu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "宁波大学",
            "name_en": "Ningbo University",
            "city": "宁波",
            "home": "https://www.nbu.edu.cn/",
            "career": "https://rsc.nbu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "商学院",
                    "unit_en": "Business School",
                    "url": "https://bs.nbu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "杭州师范大学",
            "name_en": "Hangzhou Normal University",
            "city": "杭州",
            "home": "https://www.hznu.edu.cn/",
            "career": "https://rsc.hznu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.hznu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "杭州电子科技大学",
            "name_en": "Hangzhou Dianzi University",
            "city": "杭州",
            "home": "https://www.hdu.edu.cn/",
            "career": "https://hr.hdu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://jjxy.hdu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "浙江理工大学",
            "name_en": "Zhejiang Sci-Tech University",
            "city": "杭州",
            "home": "https://www.zstu.edu.cn/",
            "career": "https://rsc.zstu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济管理学院",
                    "unit_en": "School of Economics and Management",
                    "url": "https://sem.zstu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "宁波诺丁汉大学",
            "name_en": "University of Nottingham Ningbo China",
            "city": "宁波",
            "home": "https://www.nottingham.edu.cn/",
            "career": "https://www.nottingham.edu.cn/cn/jobs/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://www.nottingham.edu.cn/cn/economics/",
                }
            ],
        },
        {
            "name_zh": "嘉兴大学",
            "name_en": "Jiaxing University",
            "city": "嘉兴",
            "home": "https://www.zjxu.edu.cn/",
            "career": "https://rsc.zjxu.edu.cn/",
            "econ_units": [],
        },
        {
            "name_zh": "温州肯恩大学",
            "name_en": "Wenzhou-Kean University",
            "city": "温州",
            "home": "https://www.wku.edu.cn/",
            "career": "https://www.wku.edu.cn/zh-hans/hr",
            "econ_units": [
                {
                    "unit_zh": "商学院",
                    "unit_en": "College of Business and Public Management",
                    "url": "https://www.wku.edu.cn/zh-hans/cbpm",
                }
            ],
        },
        {
            "name_zh": "浙江工业大学",
            "name_en": "Zhejiang University of Technology",
            "city": "杭州",
            "home": "https://www.zjut.edu.cn/",
            "career": "https://rsc.zjut.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "School of Economics",
                    "url": "https://www.zjut.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "浙江海洋大学",
            "name_en": "Zhejiang Ocean University",
            "city": "舟山",
            "home": "https://www.zjou.edu.cn/",
            "career": "https://rsc.zjou.edu.cn/",
            "econ_units": [],
        },
    ],
    "shenzhen": [
        {
            "name_zh": "深圳大学",
            "name_en": "Shenzhen University",
            "city": "深圳",
            "home": "https://www.szu.edu.cn/",
            "career": "https://hr.szu.edu.cn/",
            "apply": "http://zp.szu.edu.cn/",
            "postdoc": "https://hr.szu.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "经济学院",
                    "unit_en": "College of Economics",
                    "url": "https://eco.szu.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "南方科技大学",
            "name_en": "Southern University of Science and Technology",
            "city": "深圳",
            "home": "https://www.sustech.edu.cn/",
            "career": "https://jobs.sustech.edu.cn/",
            "postdoc": "https://jobs.sustech.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "商学院",
                    "unit_en": "School of Business",
                    "url": "https://business.sustech.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "香港中文大学（深圳）",
            "name_en": "CUHK Shenzhen",
            "city": "深圳",
            "home": "https://www.cuhk.edu.cn/",
            "career": "https://www.cuhk.edu.cn/zh-hans/taxonomy/term/38",
            "postdoc": "https://www.cuhk.edu.cn/zh-hans/taxonomy/term/50",
            "econ_units": [
                {
                    "unit_zh": "经管学院",
                    "unit_en": "SME",
                    "url": "https://sme.cuhk.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "哈尔滨工业大学（深圳）",
            "name_en": "HIT Shenzhen",
            "city": "深圳",
            "home": "https://www.hitsz.edu.cn/",
            "career": "https://hr.hitsz.edu.cn/rczp/zpxxhz.htm",
            "postdoc": "https://hr.hitsz.edu.cn/info/1009/2203.htm",
            "econ_units": [
                {
                    "unit_zh": "经济管理学院",
                    "unit_en": "School of Economics and Management",
                    "url": "https://sem.hitsz.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "深圳北理莫斯科大学",
            "name_en": "SMBU",
            "city": "深圳",
            "home": "https://www.smbu.edu.cn/",
            "career": "https://www.smbu.edu.cn/",
            "econ_units": [],
        },
        {
            "name_zh": "清华大学深圳国际研究生院",
            "name_en": "Tsinghua SIGS",
            "city": "深圳",
            "home": "https://www.sigs.tsinghua.edu.cn/",
            "career": "https://talent.sigs.tsinghua.edu.cn/index.html?locale=zh_CN",
            "postdoc": "https://talent.sigs.tsinghua.edu.cn/index.html?locale=zh_CN",
            "econ_units": [],
        },
        {
            "name_zh": "北京大学深圳研究生院",
            "name_en": "PKU Shenzhen",
            "city": "深圳",
            "home": "https://www.pkusz.edu.cn/",
            "career": "https://www.pkusz.edu.cn/",
            "econ_units": [
                {
                    "unit_zh": "汇丰商学院",
                    "unit_en": "PHBS",
                    "url": "https://phbs.pku.edu.cn/",
                }
            ],
        },
        {
            "name_zh": "中山大学深圳校区",
            "name_en": "SYSU Shenzhen",
            "city": "深圳",
            "home": "https://www.sysu.edu.cn/",
            "career": "https://rcb.sysu.edu.cn/zp",
            "econ_units": [
                {
                    "unit_zh": "岭南学院/深圳经济研究院",
                    "unit_en": "Lingnan / Shenzhen Institute of Economics",
                    "url": "https://lingnan.sysu.edu.cn/",
                }
            ],
        },
    ],
}

# Faculty + postdoc (博士后) jobs verified from Chinese career / gaoxiaojob pages
KNOWN_JOBS = [
    dict(
        region="shenzhen",
        Country="China",
        Location="Shenzhen",
        University="Shenzhen University",
        School="College of Economics",
        Deadline=None,
        JobPostTitle="Pre-tenure / tenure-track faculty (Economics, Digital Economy, etc.)",
        JobPostUrl="https://www.gaoxiaojob.com/announcement/detail/435339.html",
        RequiredMaterial="Apply via http://zp.szu.edu.cn (普通教师招聘); PhD",
        SubmissionChannel="SZU-HR",
        ApplicationLink="http://zp.szu.edu.cn/",
        Email="dyy@szu.edu.cn",
        JobType="Academic",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou / Shenzhen",
        University="Sun Yat-sen University",
        School="Lingnan College",
        Deadline=None,
        JobPostTitle="Faculty 2026 — Macro / Finance / Digital Economy / Political Economy (all fields welcome)",
        JobPostUrl="https://www.gaoxiaojob.com/announcement/detail/434822.html",
        RequiredMaterial="CV by email; rolling; Guangzhou or Shenzhen campus",
        SubmissionChannel="Email",
        ApplicationLink="https://lingnan.sysu.edu.cn/article/2580",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou / Shenzhen",
        University="Sun Yat-sen University",
        School="Lingnan College",
        Deadline=None,
        JobPostTitle="博士后 — 理论/应用经济学（含宏观、金融、数字经济、政治经济学等）",
        JobPostUrl="https://lingnan.sysu.edu.cn/article/2580",
        RequiredMaterial="年龄35周岁以下，获博士学位不超过3年；可申请逸仙博士后；优秀者可兼助理教授",
        SubmissionChannel="Email",
        ApplicationLink="https://lingnan.sysu.edu.cn/article/2580",
        Email=None,
        JobType="Postdoc",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Sun Yat-sen University",
        School="中观经济学研究院 / Lingnan",
        Deadline=None,
        JobPostTitle="博士后 — 中观经济学研究院（区域政府经济行为等）",
        JobPostUrl="https://lingnan.sysu.edu.cn/article/2774",
        RequiredMaterial="CV to sunyxi@mail.sysu.edu.cn；主题：姓名+简历+意向导师",
        SubmissionChannel="Email",
        ApplicationLink="https://lingnan.sysu.edu.cn/article/2774",
        Email="sunyxi@mail.sysu.edu.cn",
        JobType="Postdoc",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Sun Yat-sen University",
        School="国家发展研究院",
        Deadline="2026-04-15",
        JobPostTitle="博士后 — 国家发展研究院（增长/产业与数字经济/金融等）",
        JobPostUrl="https://m.gaoxiaojob.com/job/detail/2073763.html",
        RequiredMaterial="简历发意向导师；截止2026-04-15",
        SubmissionChannel="Email",
        ApplicationLink="https://m.gaoxiaojob.com/job/detail/2073763.html",
        Email=None,
        JobType="Postdoc",
    ),
    dict(
        region="shenzhen",
        Country="China",
        Location="Shenzhen",
        University="Sun Yat-sen University (Shenzhen)",
        School="Lingnan College / Shenzhen Institute of Economics",
        Deadline=None,
        JobPostTitle="Faculty 2026 — Macro / Finance / Digital Economy / Political Economy",
        JobPostUrl="https://www.gaoxiaojob.com/announcement/detail/434822.html",
        RequiredMaterial="CV by email; may choose Shenzhen campus",
        SubmissionChannel="Email",
        ApplicationLink="https://lingnan.sysu.edu.cn/",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="shenzhen",
        Country="China",
        Location="Shenzhen",
        University="Sun Yat-sen University (Shenzhen)",
        School="Lingnan College",
        Deadline=None,
        JobPostTitle="博士后 — 岭南学院（可选深圳校区）",
        JobPostUrl="https://lingnan.sysu.edu.cn/article/2580",
        RequiredMaterial="年龄35周岁以下，获博士学位不超过3年；可选广州/深圳校区",
        SubmissionChannel="Email",
        ApplicationLink="https://lingnan.sysu.edu.cn/article/2580",
        Email=None,
        JobType="Postdoc",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Jinan University",
        School="Institute for Economic and Social Research (IESR)",
        Deadline=None,
        JobPostTitle="Faculty / research positions (check IESR career page)",
        JobPostUrl="https://iesr.jnu.edu.cn/",
        RequiredMaterial="See IESR website",
        SubmissionChannel="Dept",
        ApplicationLink="https://iesr.jnu.edu.cn/",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Jinan University",
        School="经济与社会研究院 IESR / 冯帅章团队",
        Deadline=None,
        JobPostTitle="博士后 — 数字经济与AI时代儿童发展/人力资本、劳动力市场",
        JobPostUrl="https://iesr.jnu.edu.cn/2025/0624/c17209a839104/page.htm",
        RequiredMaterial="简历+代表性成果发 iesr@jnu.edu.cn；经济学相关博士，年龄≤35",
        SubmissionChannel="Email",
        ApplicationLink="https://iesr.jnu.edu.cn/2025/0624/c17209a839104/page.htm",
        Email="iesr@jnu.edu.cn",
        JobType="Postdoc",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Jinan University",
        School="大学博士后流动站（含理论/应用经济学）",
        Deadline=None,
        JobPostTitle="博士后全球招聘 2026 — 应用经济学 / 理论经济学等",
        JobPostUrl="https://hrdam.jnu.edu.cn/2026/0413/c264a853500/page.htm",
        RequiredMaterial="全年接收；与合作导师达成意向后办理进站",
        SubmissionChannel="HR/导师",
        ApplicationLink="https://hrdam.jnu.edu.cn/2026/0413/c264a853500/page.htm",
        Email=None,
        JobType="Postdoc",
    ),
    dict(
        region="guangdong",
        Country="China",
        Location="Guangzhou",
        University="Guangdong University of Foreign Studies",
        School="人力资源部 / 各学院",
        Deadline=None,
        JobPostTitle="师资博士后 2026（含经济贸易等相关学科）",
        JobPostUrl="https://rsc.gdufs.edu.cn/info/1017/12559.htm",
        RequiredMaterial="按简章填申请表；常年接受进站申请",
        SubmissionChannel="HR/学院",
        ApplicationLink="https://rsc.gdufs.edu.cn/info/1017/12559.htm",
        Email="gpszk@gdufs.edu.cn",
        JobType="Postdoc",
    ),
    dict(
        region="fujian",
        Country="China",
        Location="Xiamen",
        University="Xiamen University",
        School="School of Economics / WISE",
        Deadline=None,
        JobPostTitle="Faculty openings (check SOE / WISE / HR)",
        JobPostUrl="https://soe.xmu.edu.cn/",
        RequiredMaterial="See SOE / WISE / hr.xmu.edu.cn",
        SubmissionChannel="Dept/HR",
        ApplicationLink="https://hr.xmu.edu.cn/",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="fujian",
        Country="China",
        Location="Fuzhou",
        University="Fujian Normal University",
        School="School of Economics",
        Deadline="2026-12-31",
        JobPostTitle="Teaching-research faculty (Theory / Applied Economics) — 3 posts",
        JobPostUrl="https://www.gaoxiaojob.com/job/detail/2049266.html",
        RequiredMaterial="PhD; apply per school HR announcement",
        SubmissionChannel="HR/Gaoxiaojob",
        ApplicationLink="https://www.gaoxiaojob.com/job/detail/2049266.html",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="fujian",
        Country="China",
        Location="Fuzhou",
        University="Fujian Jiangxia University",
        School="School of Economics and Trade",
        Deadline="2026-12-31",
        JobPostTitle="Full-time faculty — Applied Economics / Digital Economy — 4 posts",
        JobPostUrl="https://www.gaoxiaojob.com/job/detail/2166774.html",
        RequiredMaterial="PhD by 2026-12-31; age limits apply",
        SubmissionChannel="HR/Gaoxiaojob",
        ApplicationLink="https://www.gaoxiaojob.com/job/detail/2166774.html",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="zhejiang",
        Country="China",
        Location="Hangzhou",
        University="Zhejiang University of Finance and Economics",
        School="University-wide (incl. Applied / Theory Economics)",
        Deadline=None,
        JobPostTitle="2026 talent recruitment — Applied Economics / Theory Economics among fields",
        JobPostUrl="http://www.shuobojob.cn/m/view.php?aid=4629",
        RequiredMaterial="Email application form to hiring unit; rolling until filled",
        SubmissionChannel="Email/HR",
        ApplicationLink="https://rsc.zufe.edu.cn/",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="zhejiang",
        Country="China",
        Location="Hangzhou",
        University="Zhejiang University",
        School="School of Economics",
        Deadline=None,
        JobPostTitle="Faculty (check School of Economics / ZJU HR)",
        JobPostUrl="http://www.cec.zju.edu.cn/",
        RequiredMaterial="See cec.zju.edu.cn / hr.zju.edu.cn",
        SubmissionChannel="Dept/HR",
        ApplicationLink="https://hr.zju.edu.cn/",
        Email=None,
        JobType="Academic",
    ),
    dict(
        region="shenzhen",
        Country="China",
        Location="Shenzhen",
        University="Chinese University of Hong Kong, Shenzhen",
        School="School of Management and Economics",
        Deadline=None,
        JobPostTitle="Faculty (check SME / CUHK-Shenzhen jobs)",
        JobPostUrl="https://sme.cuhk.edu.cn/",
        RequiredMaterial="See SME careers / university jobs page",
        SubmissionChannel="Dept",
        ApplicationLink="https://www.cuhk.edu.cn/en/page/4897",
        Email=None,
        JobType="Academic",
    ),
]


def region_dir(region: str) -> Path:
    return DATA_ECON / region


def slug(s: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", s, flags=re.U)
    return s.strip("_")[:80] or "item"


def _decode_http_body(raw: bytes, headers=None) -> str:
    """Decode bytes to text, handling gzip/deflate that some servers send anyway."""
    if not raw:
        return ""
    encoding = ""
    if headers is not None:
        try:
            encoding = (headers.get("Content-Encoding") or "").lower()
        except Exception:
            encoding = ""
    try:
        if encoding == "gzip" or raw[:2] == b"\x1f\x8b":
            raw = gzip.decompress(raw)
        elif encoding == "deflate":
            raw = zlib.decompress(raw)
    except Exception:
        pass
    for enc in ("utf-8", "gb18030", "gbk", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def fetch(url: str, timeout: int = 20) -> tuple[int | None, str]:
    if not url:
        return None, ""
    req = Request(
        url,
        headers={
            "User-Agent": UA,
            "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            # Prefer identity, but still decode gzip if the server ignores this.
            "Accept-Encoding": "gzip, deflate",
        },
    )
    try:
        with urlopen(req, timeout=timeout, context=CTX) as r:
            raw = r.read()
            return r.status, _decode_http_body(raw, r.headers)
    except HTTPError as e:
        try:
            raw = e.read()
            body = _decode_http_body(raw, e.headers)
        except Exception:
            body = ""
        return e.code, body
    except Exception as e:
        return None, str(e)


def save_fetch(folder: Path, name: str, url: str) -> dict:
    status, body = fetch(url)
    path = folder / f"{slug(name)}.html"
    if body and len(body) > 80:
        path.write_text(body, encoding="utf-8", errors="replace")
    meta = {
        "name": name,
        "url": url,
        "status": status,
        "bytes": len(body or ""),
        "file": path.name if path.exists() else None,
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
    }
    print(f"  [{status}] {name[:55]} ({meta['bytes']} B)")
    time.sleep(SLEEP)
    return meta


def parse_deadline(raw):
    if not raw:
        return None
    if isinstance(raw, datetime):
        return raw
    try:
        return datetime.strptime(str(raw)[:10], "%Y-%m-%d")
    except ValueError:
        return None


def job_to_row(job: dict) -> dict:
    url = job.get("ApplicationLink") or job.get("JobPostUrl")
    title = job.get("JobPostTitle") or "Faculty position"
    return {
        "Country": job.get("Country") or "China",
        "Location": job.get("Location"),
        "University": job.get("University"),
        "School": job.get("School"),
        "Deadline": parse_deadline(job.get("Deadline")),
        "Job Post": f"{title} | {url}" if url else title,
        "Required Material": job.get("RequiredMaterial"),
        "Submission Channel": job.get("SubmissionChannel"),
        "Application Link": url,
        "Email": job.get("Email"),
        "Submission Date": None,
        "Reference Status": None,
        "Interview Date": None,
        "Job Type": job.get("JobType") or "Academic",
        "_title": title,
        "_region": job.get("region"),
    }


def write_xlsx(rows: list[dict], path: Path, sheet: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet[:31]
    link_font = Font(color="0563C1", underline="single")
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            val = row.get(name)
            cell = ws.cell(r_i, c, val)
            if name == "Job Post" and row.get("Application Link"):
                cell.value = row.get("_title") or "Faculty"
                cell.hyperlink = str(row["Application Link"])
                cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font
    for c in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def scrape_region(region: str) -> dict:
    unis = UNIVERSITIES[region]
    base = region_dir(region)
    raw = base / "raw"
    depts = base / "departments"
    for d in (base, raw, depts):
        d.mkdir(parents=True, exist_ok=True)

    (base / f"universities_{TODAY}.json").write_text(
        json.dumps(unis, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "universities.json").write_text(
        json.dumps(unis, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    econ_units = []
    for u in unis:
        for unit in u.get("econ_units") or []:
            unit_zh = unit.get("unit_zh") or unit.get("unit") or ""
            unit_en = unit.get("unit_en") or unit.get("unit") or ""
            econ_units.append(
                {
                    "university_zh": u["name_zh"],
                    "university_en": u["name_en"],
                    "city": u.get("city"),
                    "unit_zh": unit_zh,
                    "unit_en": unit_en,
                    "url": unit.get("url"),
                    "career": unit.get("career") or u.get("career"),
                }
            )
    (base / f"economics_units_{TODAY}.json").write_text(
        json.dumps(econ_units, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "economics_units.json").write_text(
        json.dumps(econ_units, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"=== {region.upper()} ({len(unis)} universities, {len(econ_units)} econ units) ===")
    log = []
    for u in unis:
        zh = u["name_zh"]
        if u.get("career"):
            log.append(save_fetch(raw, f"招聘_{zh}", u["career"]))
        if u.get("postdoc"):
            log.append(save_fetch(raw, f"博士后_{zh}", u["postdoc"]))
        if u.get("apply"):
            log.append(save_fetch(raw, f"应聘系统_{zh}", u["apply"]))
        if u.get("home") and u.get("home") != u.get("career"):
            log.append(save_fetch(raw, f"主页_{zh}", u["home"]))
        for unit in u.get("econ_units") or []:
            if unit.get("url"):
                unit_zh = unit.get("unit_zh") or unit.get("unit") or "学院"
                log.append(save_fetch(depts, f"{zh}_{unit_zh}", unit["url"]))

    for job in KNOWN_JOBS:
        if job.get("region") != region:
            continue
        kind = "博士后" if (job.get("JobType") or "").lower() == "postdoc" else "教师"
        log.append(
            save_fetch(raw, f"{kind}_{job['JobPostTitle']}", job["JobPostUrl"])
        )

    (base / f"fetch_log_{TODAY}.json").write_text(
        json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    ok = sum(1 for m in log if m.get("status") in (200, 301, 302) and (m.get("bytes") or 0) > 500)
    fail = [m for m in log if m.get("status") not in (200, 301, 302) or (m.get("bytes") or 0) <= 500]

    rows = [job_to_row(j) for j in KNOWN_JOBS if j.get("region") == region]
    serial = []
    for r in rows:
        item = {k: v for k, v in r.items() if not k.startswith("_")}
        if isinstance(item.get("Deadline"), datetime):
            item["Deadline"] = item["Deadline"].strftime("%Y-%m-%d")
        serial.append(item)
    (base / f"listings_{TODAY}.json").write_text(
        json.dumps(serial, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    write_xlsx(rows, base / f"listings_{TODAY}.xlsx", region.title())
    write_xlsx(rows, CLEAN_ECON / f"{region}_listings_{TODAY}.xlsx", region.title())

    result = {
        "universities": len(unis),
        "econ_units": len(econ_units),
        "fetches": len(log),
        "ok_fetches": ok,
        "fail_fetches": len(fail),
        "fail_urls": [
            {"name": m["name"], "url": m["url"], "status": m["status"], "bytes": m["bytes"]}
            for m in fail
        ],
        "open_jobs": len(rows),
    }
    print(
        f"  -> ok={ok}/{len(log)} fails={len(fail)} jobs={len(rows)}"
    )
    return result


def main(regions: list[str] | None = None) -> None:
    ensure_dirs()
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    targets = regions or list(REGIONS)
    summary = {"date": TODAY, "regions": {}}
    all_rows = []
    for region in targets:
        if region not in UNIVERSITIES:
            raise SystemExit(f"Unknown region: {region}. Choose from {REGIONS}")
        summary["regions"][region] = scrape_region(region)
        all_rows.extend(
            [job_to_row(j) for j in KNOWN_JOBS if j.get("region") == region]
        )

    if set(targets) == set(REGIONS):
        write_xlsx(
            all_rows,
            CLEAN_ECON / f"china_coastal_listings_{TODAY}.xlsx",
            "China Coastal",
        )
        (DATA_ECON / f"china_coastal_summary_{TODAY}.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    else:
        tag = "_".join(targets)
        (DATA_ECON / f"china_{tag}_summary_{TODAY}.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("Done.")


if __name__ == "__main__":
    raise SystemExit(
        "DEPRECATED: use run_province_pipeline.py.\n"
        "This module remains as a helper library (fetch/job_to_row/write_xlsx)."
    )
