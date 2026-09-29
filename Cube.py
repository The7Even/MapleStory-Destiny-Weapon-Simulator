import sys
import random
import json
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QFileDialog,
    QGroupBox, QFrame, QStackedLayout, QDialog, QProgressBar, QCheckBox, QMessageBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# ---------------------------------------------------------
# 1. 등급 색상, 비용 및 천장 기준
# ---------------------------------------------------------
TIER_COLORS = {
    "레전드리": "#a8e600",
    "유니크": "#f1c40f",
    "에픽": "#a366ff",
    "레어": "#3498db"
}

TIER_ORDER = ["레어", "에픽", "유니크", "레전드리"]
TIER_BADGE_CHARS = {"레전드리": "L", "유니크": "U", "에픽": "E", "레어": "R"}

UPGRADE_RATES = {
    "레어": 0.06,
    "에픽": 0.02,
    "유니크": 0.005
}

ROLL_COSTS = {
    "윗잠": {
        "레어": 5000000,
        "에픽": 20000000,
        "유니크": 42500000,
        "레전드리": 50000000
    },
    "에디": {
        "레어": 16250000,
        "에픽": 45500000,
        "유니크": 55250000,
        "레전드리": 65000000
    }
}

CEILING_LIMITS = {
    "윗잠": {"레어": 10, "에픽": 42, "유니크": 107},
    "에디": {"레어": 62, "에픽": 152, "유니크": 214}
}

# ---------------------------------------------------------
# 2. 윗잠 전용 엑셀(data.xlsx) 공식 확률 테이블
# ---------------------------------------------------------
MAIN_PROBABILITY_TABLES = {
    '레어': [
        [
            ('STR +13', 0.061224), ('DEX +13', 0.061224), ('INT +13', 0.061224), ('LUK +13', 0.061224),
            ('최대 HP +125', 0.061224), ('최대 MP +125', 0.061224), ('공격력 +13', 0.040816), ('마력 +13', 0.040816),
            ('STR +4%', 0.061224), ('DEX +4%', 0.061224), ('INT +4%', 0.061224), ('LUK +4%', 0.061224),
            ('공격력 +4%', 0.020408), ('마력 +4%', 0.020408), ('크리티컬 확률 +5%', 0.020408), ('데미지 +4%', 0.020408),
            ('올스탯 +6', 0.040816), ('공격 시 20% 확률로 250의 HP 회복', 0.020408), ('공격 시 20% 확률로 125의 MP 회복', 0.020408),
            ('공격 시 20% 확률로 7레벨 중독효과 적용', 0.020408), ('공격 시 10% 확률로 3레벨 기절효과 적용', 0.020408),
            ('공격 시 20% 확률로 3레벨 슬로우효과 적용', 0.020408), ('공격 시 20% 확률로 4레벨 암흑효과 적용', 0.020408),
            ('공격 시 10% 확률로 3레벨 빙결효과 적용', 0.020408), ('공격 시 10% 확률로 3레벨 봉인효과 적용', 0.020408),
            ('몬스터 방어율 무시 +20%', 0.020408)
        ],
        [
            ('STR +6', 0.109091), ('DEX +6', 0.109091), ('INT +6', 0.109091), ('LUK +6', 0.109091),
            ('최대 HP +60', 0.109091), ('최대 MP +60', 0.109091), ('공격력 +6', 0.072727), ('마력 +6', 0.072727),
            ('STR +13', 0.012245), ('DEX +13', 0.012245), ('INT +13', 0.012245), ('LUK +13', 0.012245),
            ('최대 HP +125', 0.012245), ('최대 MP +125', 0.012245), ('공격력 +13', 0.008163), ('마력 +13', 0.008163),
            ('STR +4%', 0.012245), ('DEX +4%', 0.012245), ('INT +4%', 0.012245), ('LUK +4%', 0.012245),
            ('공격력 +4%', 0.004082), ('마력 +4%', 0.004082), ('크리티컬 확률 +5%', 0.004082), ('데미지 +4%', 0.004082),
            ('올스탯 +6', 0.008163), ('공격 시 20% 확률로 250의 HP 회복', 0.004082), ('공격 시 20% 확률로 125의 MP 회복', 0.004082),
            ('공격 시 20% 확률로 7레벨 중독효과 적용', 0.004082), ('공격 시 10% 확률로 3레벨 기절효과 적용', 0.004082),
            ('공격 시 20% 확률로 3레벨 슬로우효과 적용', 0.004082), ('공격 시 20% 확률로 4레벨 암흑효과 적용', 0.004082),
            ('공격 시 10% 확률로 3레벨 빙결효과 적용', 0.004082), ('공격 시 10% 확률로 3레벨 봉인효과 적용', 0.004082),
            ('몬스터 방어율 무시 +20%', 0.004082)
        ],
        [
            ('STR +6', 0.129545), ('DEX +6', 0.129545), ('INT +6', 0.129545), ('LUK +6', 0.129545),
            ('최대 HP +60', 0.129545), ('최대 MP +60', 0.129545), ('공격력 +6', 0.086364), ('마력 +6', 0.086364),
            ('STR +13', 0.003061), ('DEX +13', 0.003061), ('INT +13', 0.003061), ('LUK +13', 0.003061),
            ('최대 HP +125', 0.003061), ('최대 MP +125', 0.003061), ('공격력 +13', 0.002041), ('마력 +13', 0.002041),
            ('STR +4%', 0.003061), ('DEX +4%', 0.003061), ('INT +4%', 0.003061), ('LUK +4%', 0.003061),
            ('공격력 +4%', 0.001020), ('마력 +4%', 0.001020), ('크리티컬 확률 +5%', 0.001020), ('데미지 +4%', 0.001020),
            ('올스탯 +6', 0.002041), ('공격 시 20% 확률로 250의 HP 회복', 0.001020), ('공격 시 20% 확률로 125의 MP 회복', 0.001020),
            ('공격 시 20% 확률로 7레벨 중독효과 적용', 0.001020), ('공격 시 10% 확률로 3레벨 기절효과 적용', 0.001020),
            ('공격 시 20% 확률로 3레벨 슬로우효과 적용', 0.001020), ('공격 시 20% 확률로 4레벨 암흑효과 적용', 0.001020),
            ('공격 시 10% 확률로 3레벨 빙결효과 적용', 0.001020), ('공격 시 10% 확률로 3레벨 봉인효과 적용', 0.001020),
            ('몬스터 방어율 무시 +20%', 0.001020)
        ]
    ],
    '에픽': [
        [
            ('STR +7%', 0.108696), ('DEX +7%', 0.108696), ('INT +7%', 0.108696), ('LUK +7%', 0.108696),
            ('최대 HP +7%', 0.108696), ('최대 MP +7%', 0.108696), ('공격력 +7%', 0.043478), ('마력 +7%', 0.043478),
            ('크리티컬 확률 +9%', 0.043478), ('데미지 +7%', 0.043478), ('올스탯 +4%', 0.043478),
            ('공격 시 20% 확률로 370의 HP 회복', 0.043478), ('공격 시 20% 확률로 195의 MP 회복', 0.043478),
            ('몬스터 방어율 무시 +20%', 0.043478)
        ],
        [
            ('STR +13', 0.04898), ('DEX +13', 0.04898), ('INT +13', 0.04898), ('LUK +13', 0.04898),
            ('최대 HP +125', 0.04898), ('최대 MP +125', 0.04898), ('공격력 +13', 0.032653), ('마력 +13', 0.032653),
            ('STR +4%', 0.04898), ('DEX +4%', 0.04898), ('INT +4%', 0.04898), ('LUK +4%', 0.04898),
            ('공격력 +4%', 0.016327), ('마력 +4%', 0.016327), ('크리티컬 확률 +5%', 0.016327), ('데미지 +4%', 0.016327),
            ('올스탯 +6', 0.032653), ('공격 시 20% 확률로 250의 HP 회복', 0.016327), ('공격 시 20% 확률로 125의 MP 회복', 0.016327),
            ('공격 시 20% 확률로 7레벨 중독효과 적용', 0.016327), ('공격 시 10% 확률로 3레벨 기절효과 적용', 0.016327),
            ('공격 시 20% 확률로 3레벨 슬로우효과 적용', 0.016327), ('공격 시 20% 확률로 4레벨 암흑효과 적용', 0.016327),
            ('공격 시 10% 확률로 3레벨 빙결효과 적용', 0.016327), ('공격 시 10% 확률로 3레벨 봉인효과 적용', 0.016327),
            ('몬스터 방어율 무시 +20%', 0.016327), ('STR +7%', 0.021739), ('DEX +7%', 0.021739),
            ('INT +7%', 0.021739), ('LUK +7%', 0.021739), ('최대 HP +7%', 0.021739), ('최대 MP +7%', 0.021739),
            ('공격력 +7%', 0.008696), ('마력 +7%', 0.008696), ('크리티컬 확률 +9%', 0.008696), ('데미지 +7%', 0.008696),
            ('올스탯 +4%', 0.008696), ('공격 시 20% 확률로 370의 HP 회복', 0.008696), ('공격 시 20% 확률로 195의 MP 회복', 0.008696),
            ('몬스터 방어율 무시 +20%', 0.008696)
        ],
        [
            ('STR +13', 0.058163), ('DEX +13', 0.058163), ('INT +13', 0.058163), ('LUK +13', 0.058163),
            ('최대 HP +125', 0.058163), ('최대 MP +125', 0.058163), ('공격력 +13', 0.038776), ('마력 +13', 0.038776),
            ('STR +4%', 0.058163), ('DEX +4%', 0.058163), ('INT +4%', 0.058163), ('LUK +4%', 0.058163),
            ('공격력 +4%', 0.019388), ('마력 +4%', 0.019388), ('크리티컬 확률 +5%', 0.019388), ('데미지 +4%', 0.019388),
            ('올스탯 +6', 0.038776), ('공격 시 20% 확률로 250의 HP 회복', 0.019388), ('공격 시 20% 확률로 125의 MP 회복', 0.019388),
            ('공격 시 20% 확률로 7레벨 중독효과 적용', 0.019388), ('공격 시 10% 확률로 3레벨 기절효과 적용', 0.019388),
            ('공격 시 20% 확률로 3레벨 슬로우효과 적용', 0.019388), ('공격 시 20% 확률로 4레벨 암흑효과 적용', 0.019388),
            ('공격 시 10% 확률로 3레벨 빙결효과 적용', 0.019388), ('공격 시 10% 확률로 3레벨 봉인효과 적용', 0.019388),
            ('몬스터 방어율 무시 +20%', 0.019388), ('STR +7%', 0.005435), ('DEX +7%', 0.005435),
            ('INT +7%', 0.005435), ('LUK +7%', 0.005435), ('최대 HP +7%', 0.005435), ('최대 MP +7%', 0.005435),
            ('공격력 +7%', 0.002174), ('마력 +7%', 0.002174), ('크리티컬 확률 +9%', 0.002174), ('데미지 +7%', 0.002174),
            ('올스탯 +4%', 0.002174), ('공격 시 20% 확률로 370의 HP 회복', 0.002174), ('공격 시 20% 확률로 195의 MP 회복', 0.002174),
            ('몬스터 방어율 무시 +20%', 0.002174)
        ]
    ],
    '유니크': [
        [
            ('STR +10%', 0.116279), ('DEX +10%', 0.116279), ('INT +10%', 0.116279), ('LUK +10%', 0.116279),
            ('공격력 +10%', 0.069767), ('마력 +10%', 0.069767), ('크리티컬 확률 +10%', 0.093023), ('데미지 +10%', 0.069767),
            ('올스탯 +7%', 0.093023), ('몬스터 방어율 무시 +35%', 0.069767), ('보스 몬스터 데미지 +35%', 0.069767)
        ],
        [
            ('STR +7%', 0.086957), ('DEX +7%', 0.086957), ('INT +7%', 0.086957), ('LUK +7%', 0.086957),
            ('최대 HP +7%', 0.086957), ('최대 MP +7%', 0.086957), ('공격력 +7%', 0.034783), ('마력 +7%', 0.034783),
            ('크리티컬 확률 +9%', 0.034783), ('데미지 +7%', 0.034783), ('올스탯 +4%', 0.034783),
            ('공격 시 20% 확률로 370의 HP 회복', 0.034783), ('공격 시 20% 확률로 195의 MP 회복', 0.034783),
            ('몬스터 방어율 무시 +20%', 0.034783), ('STR +10%', 0.023256), ('DEX +10%', 0.023256),
            ('INT +10%', 0.023256), ('LUK +10%', 0.023256), ('공격력 +10%', 0.013953), ('마력 +10%', 0.013953),
            ('크리티컬 확률 +10%', 0.018605), ('데미지 +10%', 0.013953), ('올스탯 +7%', 0.018605),
            ('몬스터 방어율 무시 +35%', 0.013953), ('보스 몬스터 데미지 +35%', 0.013953)
        ],
        [
            ('STR +7%', 0.103261), ('DEX +7%', 0.103261), ('INT +7%', 0.103261), ('LUK +7%', 0.103261),
            ('최대 HP +7%', 0.103261), ('최대 MP +7%', 0.103261), ('공격력 +7%', 0.041304), ('마력 +7%', 0.041304),
            ('크리티컬 확률 +9%', 0.041304), ('데미지 +7%', 0.041304), ('올스탯 +4%', 0.041304),
            ('공격 시 20% 확률로 370의 HP 회복', 0.041304), ('공격 시 20% 확률로 195의 MP 회복', 0.041304),
            ('몬스터 방어율 무시 +20%', 0.041304), ('STR +10%', 0.005814), ('DEX +10%', 0.005814),
            ('INT +10%', 0.005814), ('LUK +10%', 0.005814), ('공격력 +10%', 0.003488), ('마력 +10%', 0.003488),
            ('크리티컬 확률 +10%', 0.004651), ('데미지 +10%', 0.003488), ('올스탯 +7%', 0.004651),
            ('몬스터 방어율 무시 +35%', 0.003488), ('보스 몬스터 데미지 +35%', 0.003488)
        ]
    ],
    '레전드리': [
        [
            ('STR +13%', 0.097561), ('DEX +13%', 0.097561), ('INT +13%', 0.097561), ('LUK +13%', 0.097561),
            ('공격력 +13%', 0.048780), ('마력 +13%', 0.048780), ('크리티컬 확률 +13%', 0.048780), ('데미지 +13%', 0.048780),
            ('올스탯 +10%', 0.073171), ('공격력 +32', 0.048780), ('마력 +32', 0.048780), ('몬스터 방어율 무시 +40%', 0.048780),
            ('몬스터 방어율 무시 +45%', 0.048780), ('보스 몬스터 데미지 +40%', 0.097561), ('보스 몬스터 데미지 +45%', 0.048780)
        ],
        [
            ('STR +10%', 0.093023), ('DEX +10%', 0.093023), ('INT +10%', 0.093023), ('LUK +10%', 0.093023),
            ('공격력 +10%', 0.055814), ('마력 +10%', 0.055814), ('크리티컬 확률 +10%', 0.074419), ('데미지 +10%', 0.055814),
            ('올스탯 +7%', 0.074419), ('몬스터 방어율 무시 +35%', 0.055814), ('보스 몬스터 데미지 +35%', 0.055814),
            ('STR +13%', 0.019512), ('DEX +13%', 0.019512), ('INT +13%', 0.019512), ('LUK +13%', 0.019512),
            ('공격력 +13%', 0.009756), ('마력 +13%', 0.009756), ('크리티컬 확률 +13%', 0.009756), ('데미지 +13%', 0.009756),
            ('올스탯 +10%', 0.014634), ('공격력 +32', 0.009756), ('마력 +32', 0.009756), ('몬스터 방어율 무시 +40%', 0.009756),
            ('몬스터 방어율 무시 +45%', 0.009756), ('보스 몬스터 데미지 +40%', 0.019512), ('보스 몬스터 데미지 +45%', 0.009756)
        ],
        [
            ('STR +10%', 0.110465), ('DEX +10%', 0.110465), ('INT +10%', 0.110465), ('LUK +10%', 0.110465),
            ('공격력 +10%', 0.066279), ('마력 +10%', 0.066279), ('크리티컬 확률 +10%', 0.088372), ('데미지 +10%', 0.066279),
            ('올스탯 +7%', 0.088372), ('몬스터 방어율 무시 +35%', 0.066279), ('보스 몬스터 데미지 +35%', 0.066279),
            ('STR +13%', 0.004878), ('DEX +13%', 0.004878), ('INT +13%', 0.004878), ('LUK +13%', 0.004878),
            ('공격력 +13%', 0.002439), ('마력 +13%', 0.002439), ('크리티컬 확률 +13%', 0.002439), ('데미지 +13%', 0.002439),
            ('올스탯 +10%', 0.003659), ('공격력 +32', 0.002439), ('마력 +32', 0.002439), ('몬스터 방어율 무시 +40%', 0.002439),
            ('몬스터 방어율 무시 +45%', 0.002439), ('보스 몬스터 데미지 +40%', 0.004878), ('보스 몬스터 데미지 +45%', 0.002439)
        ]
    ]
}

MAIN_TIER_SLOT1_SETS = {
    tier: set(opt for opt, _ in MAIN_PROBABILITY_TABLES[tier][0])
    for tier in TIER_ORDER
}

# ---------------------------------------------------------
# 3. 에디셔널 전용 엑셀(data_2.xlsx / E.시트) 공식 확률 테이블
# ---------------------------------------------------------
ADD_PROBABILITY_TABLES = {
    '레어': [
        [
            ('최대 HP +125', 0.058824), ('최대 MP +125', 0.058824), ('이동속도 +6', 0.058824), ('점프력 +6', 0.058824),
            ('방어력 +125', 0.058824), ('STR +13', 0.058824), ('DEX +13', 0.058824), ('INT +13', 0.058824),
            ('LUK +13', 0.058824), ('공격력 +13', 0.039216), ('마력 +13', 0.039216), ('최대 HP +3%', 0.039216),
            ('최대 MP +3%', 0.039216), ('STR +4%', 0.039216), ('DEX +4%', 0.039216), ('INT +4%', 0.039216),
            ('LUK +4%', 0.039216), ('공격력 +4%', 0.019608), ('마력 +4%', 0.019608), ('크리티컬 확률 +5%', 0.039216),
            ('데미지 +4%', 0.019608), ('올스탯 +6', 0.058824)
        ],
        [
            ('STR +6', 0.072622), ('DEX +6', 0.072622), ('INT +6', 0.072622), ('LUK +6', 0.072622),
            ('최대 HP +60', 0.108932), ('최대 MP +60', 0.108932), ('이동속도 +4', 0.108932), ('점프력 +4', 0.108932),
            ('방어력 +60', 0.108932), ('공격력 +6', 0.072622), ('마력 +6', 0.072622), ('최대 HP +125', 0.001153),
            ('최대 MP +125', 0.001153), ('이동속도 +6', 0.001153), ('점프력 +6', 0.001153), ('방어력 +125', 0.001153),
            ('STR +13', 0.001153), ('DEX +13', 0.001153), ('INT +13', 0.001153), ('LUK +13', 0.001153),
            ('공격력 +13', 0.000769), ('마력 +13', 0.000769), ('최대 HP +3%', 0.000769), ('최대 MP +3%', 0.000769),
            ('STR +4%', 0.000769), ('DEX +4%', 0.000769), ('INT +4%', 0.000769), ('LUK +4%', 0.000769),
            ('공격력 +4%', 0.000385), ('마력 +4%', 0.000385), ('크리티컬 확률 +5%', 0.000769), ('데미지 +4%', 0.000385),
            ('올스탯 +6', 0.001153)
        ],
        [
            ('STR +6', 0.072622), ('DEX +6', 0.072622), ('INT +6', 0.072622), ('LUK +6', 0.072622),
            ('최대 HP +60', 0.108932), ('최대 MP +60', 0.108932), ('이동속도 +4', 0.108932), ('점프력 +4', 0.108932),
            ('방어력 +60', 0.108932), ('공격력 +6', 0.072622), ('마력 +6', 0.072622), ('최대 HP +125', 0.001153),
            ('최대 MP +125', 0.001153), ('이동속도 +6', 0.001153), ('점프력 +6', 0.001153), ('방어력 +125', 0.001153),
            ('STR +13', 0.001153), ('DEX +13', 0.001153), ('INT +13', 0.001153), ('LUK +13', 0.001153),
            ('공격력 +13', 0.000769), ('마력 +13', 0.000769), ('최대 HP +3%', 0.000769), ('최대 MP +3%', 0.000769),
            ('STR +4%', 0.000769), ('DEX +4%', 0.000769), ('INT +4%', 0.000769), ('LUK +4%', 0.000769),
            ('공격력 +4%', 0.000385), ('마력 +4%', 0.000385), ('크리티컬 확률 +5%', 0.000769), ('데미지 +4%', 0.000385),
            ('올스탯 +6', 0.001153)
        ]
    ],
    '에픽': [
        [
            ('최대 HP +6%', 0.088235), ('최대 MP +6%', 0.088235), ('공격력 +7%', 0.058824), ('마력 +7%', 0.058824),
            ('크리티컬 확률 +7%', 0.029412), ('STR +7%', 0.088235), ('DEX +7%', 0.088235), ('INT +7%', 0.088235),
            ('LUK +7%', 0.088235), ('데미지 +7%', 0.029412), ('올스탯 +4%', 0.058824),
            ('공격 시 3% 확률로 54의 HP 회복', 0.088235), ('공격 시 3% 확률로 54의 MP 회복', 0.088235),
            ('몬스터 방어율 무시 +5%', 0.058824)
        ],
        [
            ('최대 HP +125', 0.056022), ('최대 MP +125', 0.056022), ('이동속도 +6', 0.056022), ('점프력 +6', 0.056022),
            ('방어력 +125', 0.056022), ('STR +13', 0.056022), ('DEX +13', 0.056022), ('INT +13', 0.056022),
            ('LUK +13', 0.056022), ('공격력 +13', 0.037348), ('마력 +13', 0.037348), ('최대 HP +3%', 0.037348),
            ('최대 MP +3%', 0.037348), ('STR +4%', 0.037348), ('DEX +4%', 0.037348), ('INT +4%', 0.037348),
            ('LUK +4%', 0.037348), ('공격력 +4%', 0.018674), ('마력 +4%', 0.018674), ('크리티컬 확률 +5%', 0.037348),
            ('데미지 +4%', 0.018674), ('올스탯 +6', 0.056022), ('최대 HP +6%', 0.004202), ('최대 MP +6%', 0.004202),
            ('공격력 +7%', 0.002801), ('마력 +7%', 0.002801), ('크리티컬 확률 +7%', 0.001401), ('STR +7%', 0.004202),
            ('DEX +7%', 0.004202), ('INT +7%', 0.004202), ('LUK +7%', 0.004202), ('데미지 +7%', 0.001401),
            ('올스탯 +4%', 0.002801), ('공격 시 3% 확률로 54의 HP 회복', 0.004202), ('공격 시 3% 확률로 54의 MP 회복', 0.004202),
            ('몬스터 방어율 무시 +5%', 0.002801)
        ],
        [
            ('최대 HP +125', 0.056022), ('최대 MP +125', 0.056022), ('이동속도 +6', 0.056022), ('점프력 +6', 0.056022),
            ('방어력 +125', 0.056022), ('STR +13', 0.056022), ('DEX +13', 0.056022), ('INT +13', 0.056022),
            ('LUK +13', 0.056022), ('공격력 +13', 0.037348), ('마력 +13', 0.037348), ('최대 HP +3%', 0.037348),
            ('최대 MP +3%', 0.037348), ('STR +4%', 0.037348), ('DEX +4%', 0.037348), ('INT +4%', 0.037348),
            ('LUK +4%', 0.037348), ('공격력 +4%', 0.018674), ('마력 +4%', 0.018674), ('크리티컬 확률 +5%', 0.037348),
            ('데미지 +4%', 0.018674), ('올스탯 +6', 0.056022), ('최대 HP +6%', 0.004202), ('최대 MP +6%', 0.004202),
            ('공격력 +7%', 0.002801), ('마력 +7%', 0.002801), ('크리티컬 확률 +7%', 0.001401), ('STR +7%', 0.004202),
            ('DEX +7%', 0.004202), ('INT +7%', 0.004202), ('LUK +7%', 0.004202), ('데미지 +7%', 0.001401),
            ('올스탯 +4%', 0.002801), ('공격 시 3% 확률로 54의 HP 회복', 0.004202), ('공격 시 3% 확률로 54의 MP 회복', 0.004202),
            ('몬스터 방어율 무시 +5%', 0.002801)
        ]
    ],
    '유니크': [
        [
            ('최대 HP +9%', 0.069767), ('최대 MP +9%', 0.069767), ('공격력 +10%', 0.046512), ('마력 +10%', 0.046512),
            ('크리티컬 확률 +10%', 0.046512), ('STR +10%', 0.069767), ('DEX +10%', 0.069767), ('INT +10%', 0.069767),
            ('LUK +10%', 0.069767), ('데미지 +10%', 0.023256), ('올스탯 +7%', 0.046512),
            ('캐릭터 기준 9레벨 당 STR +1', 0.046512), ('캐릭터 기준 9레벨 당 DEX +1', 0.046512),
            ('캐릭터 기준 9레벨 당 INT +1', 0.046512), ('캐릭터 기준 9레벨 당 LUK +1', 0.046512),
            ('공격 시 15% 확률로 100의 HP 회복', 0.069767), ('공격 시 15% 확률로 100의 MP 회복', 0.069767),
            ('몬스터 방어율 무시 +6%', 0.023256), ('보스 몬스터 데미지 +14%', 0.023256)
        ],
        [
            ('최대 HP +6%', 0.086505), ('최대 MP +6%', 0.086505), ('공격력 +7%', 0.05767), ('마력 +7%', 0.05767),
            ('크리티컬 확률 +7%', 0.028835), ('STR +7%', 0.086505), ('DEX +7%', 0.086505), ('INT +7%', 0.086505),
            ('LUK +7%', 0.086505), ('데미지 +7%', 0.028835), ('올스탯 +4%', 0.05767),
            ('공격 시 3% 확률로 54의 HP 회복', 0.086505), ('공격 시 3% 확률로 54의 MP 회복', 0.086505),
            ('몬스터 방어율 무시 +5%', 0.05767), ('최대 HP +9%', 0.001368), ('최대 MP +9%', 0.001368),
            ('공격력 +10%', 0.000912), ('마력 +10%', 0.000912), ('크리티컬 확률 +10%', 0.000912),
            ('STR +10%', 0.001368), ('DEX +10%', 0.001368), ('INT +10%', 0.001368), ('LUK +10%', 0.001368),
            ('데미지 +10%', 0.000456), ('올스탯 +7%', 0.000912), ('캐릭터 기준 9레벨 당 STR +1', 0.000912),
            ('캐릭터 기준 9레벨 당 DEX +1', 0.000912), ('캐릭터 기준 9레벨 당 INT +1', 0.000912),
            ('캐릭터 기준 9레벨 당 LUK +1', 0.000912), ('공격 시 15% 확률로 100의 HP 회복', 0.001368),
            ('공격 시 15% 확률로 100의 MP 회복', 0.001368), ('몬스터 방어율 무시 +6%', 0.000456),
            ('보스 몬스터 데미지 +14%', 0.000456)
        ],
        [
            ('최대 HP +6%', 0.086505), ('최대 MP +6%', 0.086505), ('공격력 +7%', 0.05767), ('마력 +7%', 0.05767),
            ('크리티컬 확률 +7%', 0.028835), ('STR +7%', 0.086505), ('DEX +7%', 0.086505), ('INT +7%', 0.086505),
            ('LUK +7%', 0.086505), ('데미지 +7%', 0.028835), ('올스탯 +4%', 0.05767),
            ('공격 시 3% 확률로 54의 HP 회복', 0.086505), ('공격 시 3% 확률로 54의 MP 회복', 0.086505),
            ('몬스터 방어율 무시 +5%', 0.05767), ('최대 HP +9%', 0.001368), ('최대 MP +9%', 0.001368),
            ('공격력 +10%', 0.000912), ('마력 +10%', 0.000912), ('크리티컬 확률 +10%', 0.000912),
            ('STR +10%', 0.001368), ('DEX +10%', 0.001368), ('INT +10%', 0.001368), ('LUK +10%', 0.001368),
            ('데미지 +10%', 0.000456), ('올스탯 +7%', 0.000912), ('캐릭터 기준 9레벨 당 STR +1', 0.000912),
            ('캐릭터 기준 9레벨 당 DEX +1', 0.000912), ('캐릭터 기준 9레벨 당 INT +1', 0.000912),
            ('캐릭터 기준 9레벨 당 LUK +1', 0.000912), ('공격 시 15% 확률로 100의 HP 회복', 0.001368),
            ('공격 시 15% 확률로 100의 MP 회복', 0.001368), ('몬스터 방어율 무시 +6%', 0.000456),
            ('보스 몬스터 데미지 +14%', 0.000456)
        ]
    ],
    '레전드리': [
        [
            ('최대 HP +12%', 0.076923), ('최대 MP +12%', 0.076923), ('공격력 +13%', 0.051282), ('마력 +13%', 0.051282),
            ('크리티컬 확률 +13%', 0.051282), ('STR +13%', 0.076923), ('DEX +13%', 0.076923), ('INT +13%', 0.076923),
            ('LUK +13%', 0.076923), ('데미지 +13%', 0.025641), ('올스탯 +10%', 0.051282),
            ('캐릭터 기준 9레벨 당 STR +2', 0.051282), ('캐릭터 기준 9레벨 당 DEX +2', 0.051282),
            ('캐릭터 기준 9레벨 당 INT +2', 0.051282), ('캐릭터 기준 9레벨 당 LUK +2', 0.051282),
            ('공격력 +32', 0.025641), ('마력 +32', 0.025641), ('몬스터 방어율 무시 +7%', 0.025641),
            ('보스 몬스터 데미지 +20%', 0.025641)
        ],
        [
            ('최대 HP +9%', 0.06942), ('최대 MP +9%', 0.06942), ('공격력 +10%', 0.04628), ('마력 +10%', 0.04628),
            ('크리티컬 확률 +10%', 0.04628), ('STR +10%', 0.06942), ('DEX +10%', 0.06942), ('INT +10%', 0.06942),
            ('LUK +10%', 0.06942), ('데미지 +10%', 0.02314), ('올스탯 +7%', 0.04628),
            ('캐릭터 기준 9레벨 당 STR +1', 0.04628), ('캐릭터 기준 9레벨 당 DEX +1', 0.04628),
            ('캐릭터 기준 9레벨 당 INT +1', 0.04628), ('캐릭터 기준 9레벨 당 LUK +1', 0.04628),
            ('공격 시 15% 확률로 100의 HP 회복', 0.06942), ('공격 시 15% 확률로 100의 MP 회복', 0.06942),
            ('몬스터 방어율 무시 +6%', 0.02314), ('보스 몬스터 데미지 +14%', 0.02314),
            ('최대 HP +12%', 0.000383), ('최대 MP +12%', 0.000383), ('공격력 +13%', 0.000255), ('마력 +13%', 0.000255),
            ('크리티컬 확률 +13%', 0.000255), ('STR +13%', 0.000383), ('DEX +13%', 0.000383), ('INT +13%', 0.000383),
            ('LUK +13%', 0.000383), ('데미지 +13%', 0.000128), ('올스탯 +10%', 0.000255),
            ('캐릭터 기준 9레벨 당 STR +2', 0.000255), ('캐릭터 기준 9레벨 당 DEX +2', 0.000255),
            ('캐릭터 기준 9레벨 당 INT +2', 0.000255), ('캐릭터 기준 9레벨 당 LUK +2', 0.000255),
            ('공격력 +32', 0.000128), ('마력 +32', 0.000128), ('몬스터 방어율 무시 +7%', 0.000128),
            ('보스 몬스터 데미지 +20%', 0.000128)
        ],
        [
            ('최대 HP +9%', 0.06942), ('최대 MP +9%', 0.06942), ('공격력 +10%', 0.04628), ('마력 +10%', 0.04628),
            ('크리티컬 확률 +10%', 0.04628), ('STR +10%', 0.06942), ('DEX +10%', 0.06942), ('INT +10%', 0.06942),
            ('LUK +10%', 0.06942), ('데미지 +10%', 0.02314), ('올스탯 +7%', 0.04628),
            ('캐릭터 기준 9레벨 당 STR +1', 0.04628), ('캐릭터 기준 9레벨 당 DEX +1', 0.04628),
            ('캐릭터 기준 9레벨 당 INT +1', 0.04628), ('캐릭터 기준 9레벨 당 LUK +1', 0.04628),
            ('공격 시 15% 확률로 100의 HP 회복', 0.06942), ('공격 시 15% 확률로 100의 MP 회복', 0.06942),
            ('몬스터 방어율 무시 +6%', 0.02314), ('보스 몬스터 데미지 +14%', 0.02314),
            ('최대 HP +12%', 0.000383), ('최대 MP +12%', 0.000383), ('공격력 +13%', 0.000255), ('마력 +13%', 0.000255),
            ('크리티컬 확률 +13%', 0.000255), ('STR +13%', 0.000383), ('DEX +13%', 0.000383), ('INT +13%', 0.000383),
            ('LUK +13%', 0.000383), ('데미지 +13%', 0.000128), ('올스탯 +10%', 0.000255),
            ('캐릭터 기준 9레벨 당 STR +2', 0.000255), ('캐릭터 기준 9레벨 당 DEX +2', 0.000255),
            ('캐릭터 기준 9레벨 당 INT +2', 0.000255), ('캐릭터 기준 9레벨 당 LUK +2', 0.000255),
            ('공격력 +32', 0.000128), ('마력 +32', 0.000128), ('몬스터 방어율 무시 +7%', 0.000128),
            ('보스 몬스터 데미지 +20%', 0.000128)
        ]
    ]
}

ADD_TIER_SLOT1_SETS = {
    tier: set(opt for opt, _ in ADD_PROBABILITY_TABLES[tier][0])
    for tier in TIER_ORDER
}

# ---------------------------------------------------------
# 4. 정밀 가치 점수 환산 로직 (보공 = 뎀 동등화 보정 적용)
# ---------------------------------------------------------
def calculate_score(option_text, is_magic=False, is_additional=False):
    main_stat = "INT" if is_magic else "STR"
    sub_stat = "LUK" if is_magic else "DEX"
    att_stat = "마력" if is_magic else "공격력"

    # 1. 공/마 %: 공마 1% = 주스탯 4% = 4.0점
    if f"{att_stat} +" in option_text and "%" in option_text:
        try:
            val = float(re.search(rf"{att_stat} \+(\d+)%", option_text).group(1))
            return val * 4.0
        except:
            pass

    # 2. 보스 몬스터 데미지: 보공 45% = 52.0점, 보공 40% = 46.0점, 보공 35% = 40.0점
    if "보스 몬스터 데미지 +" in option_text:
        if "보스 몬스터 데미지 +45%" in option_text:
            return 52.0
        elif "보스 몬스터 데미지 +40%" in option_text:
            return 46.0
        elif "보스 몬스터 데미지 +35%" in option_text:
            return 40.0
        elif "보스 몬스터 데미지 +20%" in option_text:
            return 23.0
        elif "보스 몬스터 데미지 +14%" in option_text:
            return 16.0

    # 3. 일반 데미지: 뎀 1% = 보공 1% (40% 기준 46.0점 -> 1%당 1.15점)
    if "데미지 +" in option_text and "보스" not in option_text:
        try:
            val = float(re.search(r"데미지 \+(\d+)%", option_text).group(1))
            return val * 1.15
        except:
            pass

    # 4. 몬스터 방어율 무시: 방무 1% = 공마 0.1% = 0.4점
    if "몬스터 방어율 무시 +" in option_text:
        try:
            val = float(re.search(r"몬스터 방어율 무시 \+(\d+)%", option_text).group(1))
            return val * 0.4
        except:
            pass

    # 5. 올스탯 %: 주스탯 1.0 + 부스탯 0.1 = 1.1점
    if "올스탯 +" in option_text and "%" in option_text:
        try:
            val = float(re.search(r"올스탯 \+(\d+)%", option_text).group(1))
            return val * 1.1
        except:
            pass

    # 6. 주스탯 %: 1.0점
    if f"{main_stat} +" in option_text and "%" in option_text:
        try:
            val = float(re.search(rf"{main_stat} \+(\d+)%", option_text).group(1))
            return val * 1.0
        except:
            pass

    # 7. 부스탯 %: 부스탯 40% = 주스탯 4% => 부스탯 1% = 0.1점
    if f"{sub_stat} +" in option_text and "%" in option_text:
        try:
            val = float(re.search(rf"{sub_stat} \+(\d+)%", option_text).group(1))
            return val * 0.1
        except:
            pass

    # 8. 깡공/마 (플랫): 공마 +11 = 주스탯 4% => 공마 +1 = 4/11 점 ≈ 0.3636점
    if f"{att_stat} +" in option_text and "%" not in option_text:
        try:
            val = float(re.search(rf"{att_stat} \+(\d+)", option_text).group(1))
            return val * (4.0 / 11.0)
        except:
            pass

    # 9. 깡주스탯 (플랫): 주스탯 +44 = 주스탯 4% => 주스탯 +1 = 1/11 점 ≈ 0.0909점
    if f"{main_stat} +" in option_text and "%" not in option_text and "레벨" not in option_text:
        try:
            val = float(re.search(rf"{main_stat} \+(\d+)", option_text).group(1))
            return val * (1.0 / 11.0)
        except:
            pass

    # 10. 깡부스탯 (플랫): 부스탯 +440 = 주스탯 4% => 부스탯 +1 = 1/110 점 ≈ 0.0091점
    if f"{sub_stat} +" in option_text and "%" not in option_text and "레벨" not in option_text:
        try:
            val = float(re.search(rf"{sub_stat} \+(\d+)", option_text).group(1))
            return val * (0.1 / 11.0)
        except:
            pass

    # 11. 깡올스탯 (플랫): 올스탯 +1 = 주스탯 1/11 + 부스탯 1/110 = 11/110 = 0.1점
    if "올스탯 +" in option_text and "%" not in option_text:
        try:
            val = float(re.search(r"올스탯 \+(\d+)", option_text).group(1))
            return val * 0.1
        except:
            pass

    # 12. 캐릭터 기준 9레벨 당 스탯 (250제 기준 250 // 9 = 27단계)
    if "캐릭터 기준 9레벨 당" in option_text:
        try:
            val = float(re.search(r"\+(\d+)", option_text).group(1))
            total_flat = 27 * val
            if main_stat in option_text:
                return total_flat * (1.0 / 11.0)
            elif sub_stat in option_text:
                return total_flat * (0.1 / 11.0)
        except:
            pass

    return 0.0

# ---------------------------------------------------------
# 5. 모드별 맞춤 목표 드롭다운 목록
# ---------------------------------------------------------
MAIN_TARGET_OPTIONS = [
    "선택 안 함 (무관)",
    "공/마 13% 이상",
    "공/마 10% 이상",
    "보공 45%",
    "보공 40% 이상",
    "보공 35% 이상",
    "방무 40% 이상",
    "방무 35% 이상"
]

ADD_TARGET_OPTIONS = [
    "선택 안 함 (무관)",
    "공/마 13% 이상",
    "공/마 10% 이상",
    "보공 20% (레전)",
    "보공 14% 이상",
    "공/마 +32 이상",
    "방무 7% 이상"
]

def matches_target_condition(option_text, target_condition, is_magic=False):
    target_stat = "마력" if is_magic else "공격력"
    if target_condition == "선택 안 함 (무관)":
        return True
    elif target_condition == "공/마 13% 이상":
        return f"{target_stat} +13%" in option_text
    elif target_condition == "공/마 10% 이상":
        return (f"{target_stat} +10%" in option_text) or (f"{target_stat} +13%" in option_text)
    elif target_condition == "보공 45%":
        return "보스 몬스터 데미지 +45%" in option_text
    elif target_condition == "보공 40% 이상":
        return ("보스 몬스터 데미지 +40%" in option_text) or ("보스 몬스터 데미지 +45%" in option_text)
    elif target_condition == "보공 35% 이상":
        return ("보스 몬스터 데미지 +35%" in option_text) or ("보스 몬스터 데미지 +40%" in option_text) or ("보스 몬스터 데미지 +45%" in option_text)
    elif target_condition == "보공 20% (레전)":
        return "보스 몬스터 데미지 +20%" in option_text
    elif target_condition == "보공 14% 이상":
        return ("보스 몬스터 데미지 +14%" in option_text) or ("보스 몬스터 데미지 +20%" in option_text)
    elif target_condition == "공/마 +32 이상":
        return f"{target_stat} +32" in option_text
    elif target_condition == "방무 40% 이상":
        return ("몬스터 방어율 무시 +40%" in option_text) or ("몬스터 방어율 무시 +45%" in option_text)
    elif target_condition == "방무 35% 이상":
        return ("몬스터 방어율 무시 +35%" in option_text) or ("몬스터 방어율 무시 +40%" in option_text) or ("몬스터 방어율 무시 +45%" in option_text)
    elif target_condition == "방무 7% 이상":
        return ("몬스터 방어율 무시 +7%" in option_text)
    return False

def check_card_satisfaction(lines, target_conditions, stop_if_better, before_score, is_magic, is_additional):
    card_score = sum(calculate_score(l[0], is_magic, is_additional) for l in lines)
    if stop_if_better and (card_score > before_score):
        return True

    active_reqs = [req for req in target_conditions if req != "선택 안 함 (무관)"]
    if not active_reqs:
        return False

    line_texts = [l[0] for l in lines]
    matched_indices = set()

    for req in active_reqs:
        found = False
        for i, text in enumerate(line_texts):
            if i not in matched_indices and matches_target_condition(text, req, is_magic):
                matched_indices.add(i)
                found = True
                break
        if not found:
            return False

    return True

# ---------------------------------------------------------
# 6. 모드별 확률 추첨 함수
# ---------------------------------------------------------
def roll_three_lines(mode, current_tier):
    tables = MAIN_PROBABILITY_TABLES if mode == "윗잠" else ADD_PROBABILITY_TABLES
    slot1_set = MAIN_TIER_SLOT1_SETS[current_tier] if mode == "윗잠" else ADD_TIER_SLOT1_SETS[current_tier]
    tier_table = tables[current_tier]

    lines = []
    for slot_idx in range(3):
        slot_options = [opt for opt, _ in tier_table[slot_idx]]
        slot_weights = [w for _, w in tier_table[slot_idx]]
        chosen_opt = random.choices(slot_options, weights=slot_weights, k=1)[0]

        if slot_idx == 0:
            badge = TIER_BADGE_CHARS[current_tier]
        elif chosen_opt in slot1_set:
            badge = TIER_BADGE_CHARS[current_tier]
        else:
            idx = TIER_ORDER.index(current_tier)
            badge = TIER_BADGE_CHARS[TIER_ORDER[max(0, idx - 1)]]

        lines.append((chosen_opt, badge))
    return lines

# ---------------------------------------------------------
# 7. 비동기 자동 롤링 워커
# ---------------------------------------------------------
class AutoRollerThread(QThread):
    roll_tick = pyqtSignal(dict)
    finished_roll = pyqtSignal(dict)

    def __init__(self, mode, current_tier, pity_count, target_conditions, stop_if_better, before_score, is_magic):
        super().__init__()
        self.mode = mode
        self.current_tier = current_tier
        self.pity_count = pity_count
        self.target_conditions = target_conditions
        self.stop_if_better = stop_if_better
        self.before_score = before_score
        self.is_magic = is_magic
        self.running = True

    def run(self):
        tier = self.current_tier
        pity = self.pity_count
        cost_per_roll = ROLL_COSTS[self.mode][tier] * 3
        is_add = (self.mode == "에디")

        while self.running:
            candidates = []
            for _ in range(3):
                pity += 1
                lines = roll_three_lines(self.mode, tier)
                score = sum(calculate_score(line[0], self.is_magic, is_add) for line in lines)
                candidates.append((tier, lines, score))

            data = {
                "tier": tier,
                "candidates": candidates,
                "cost": cost_per_roll,
                "pity": pity,
                "rolls_count": 3
            }
            self.roll_tick.emit(data)

            satisfied_idx = -1
            for idx, cand in enumerate(candidates):
                if check_card_satisfaction(cand[1], self.target_conditions, self.stop_if_better, self.before_score, self.is_magic, is_add):
                    satisfied_idx = idx
                    break

            if satisfied_idx != -1:
                data["satisfied_idx"] = satisfied_idx
                self.finished_roll.emit(data)
                break

            self.msleep(20)

    def stop(self):
        self.running = False

# ---------------------------------------------------------
# 8. 개별 카드 컴포넌트
# ---------------------------------------------------------
class PotentialCard(QFrame):
    def __init__(self, title="AFTER", is_before=False, select_callback=None, index=0):
        super().__init__()
        self.title = title
        self.is_before = is_before
        self.select_callback = select_callback
        self.index = index
        self.tier = "레어"
        self.lines = []
        self.score = 0.0

        self.setObjectName("card_frame")
        self.setFixedSize(190, 210)
        self.setCursor(Qt.CursorShape.PointingHandCursor if not is_before else Qt.CursorShape.ArrowCursor)
        self.init_ui()

    def init_ui(self):
        self.card_stack = QStackedLayout()
        self.card_stack.setContentsMargins(0, 0, 0, 0)

        self.view_normal = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(5)

        self.lbl_title = QLabel(self.title)
        self.lbl_title.setStyleSheet("font-weight: bold; color: #a0aab5; font-size: 11px;")
        layout.addWidget(self.lbl_title)

        self.lbl_tier = QLabel("-")
        self.lbl_tier.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_tier.setFixedHeight(18)
        self.lbl_tier.setStyleSheet("background-color: #2c3e50; color: #888; font-weight: bold; border-radius: 3px; font-size: 11px;")
        layout.addWidget(self.lbl_tier)

        self.line_labels = []
        for _ in range(3):
            lbl = QLabel("")
            lbl.setStyleSheet("color: #666; font-size: 11px;")
            self.line_labels.append(lbl)
            layout.addWidget(lbl)

        layout.addStretch()

        self.diff_btn = QPushButton("-")
        self.diff_btn.setObjectName("diff_button")
        if not self.is_before and self.select_callback:
            self.diff_btn.clicked.connect(lambda: self.select_callback(self.index))
        layout.addWidget(self.diff_btn)

        self.view_normal.setLayout(layout)
        self.card_stack.addWidget(self.view_normal)

        # 물음표 단독 뷰
        self.blind_view = QWidget()
        blind_layout = QVBoxLayout()
        blind_layout.setContentsMargins(0, 0, 0, 0)
        blind_layout.setSpacing(0)

        self.btn_blind = QPushButton("?")
        self.btn_blind.setObjectName("blind_single_btn")
        self.btn_blind.setFixedSize(190, 210)
        self.btn_blind.clicked.connect(self.reveal_this_card)
        blind_layout.addWidget(self.btn_blind)

        self.blind_view.setLayout(blind_layout)
        self.card_stack.addWidget(self.blind_view)

        self.setLayout(self.card_stack)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if not self.is_before and self.select_callback and self.card_stack.currentIndex() == 0:
                self.select_callback(self.index)
        super().mousePressEvent(event)

    def show_blind(self):
        self.card_stack.setCurrentIndex(1)

    def reveal_this_card(self):
        self.card_stack.setCurrentIndex(0)

    def set_data(self, tier, lines, is_magic=False, base_score=0.0, is_upgrade=False, is_additional=False):
        self.tier = tier
        self.lines = lines
        self.diff_btn.setEnabled(True)
        self.card_stack.setCurrentIndex(0)

        color = TIER_COLORS.get(self.tier, "#3498db")
        self.lbl_tier.setText(self.tier)
        self.lbl_tier.setStyleSheet(f"background-color: {color}; color: #111; font-weight: bold; border-radius: 3px; font-size: 11px;")

        self.score = sum(calculate_score(l[0], is_magic, is_additional) for l in lines)
        for i in range(3):
            text, badge = lines[i]
            l_color = TIER_COLORS.get("레전드리" if badge == "L" else "유니크" if badge == "U" else "에픽" if badge == "E" else "레어", "#3498db")
            self.line_labels[i].setText(f"■ [{badge}] {text}")
            self.line_labels[i].setStyleSheet(f"color: {l_color}; font-size: 11px; font-weight: bold;")

        if self.is_before:
            self.diff_btn.setText(f"현재 환산 스탯\n({self.score:.1f}점)")
            self.diff_btn.setStyleSheet("background-color: #242c38; color: #8a95a5;")
        elif is_upgrade:
            self.diff_btn.setText("★ 등급 상승! 클릭 확정 ★")
            self.diff_btn.setStyleSheet("""
                background-color: #e67e22;
                color: #ffffff;
                font-weight: bold;
                border: 2px solid #f1c40f;
                padding: 6px;
            """)
        else:
            diff = self.score - base_score
            sign = "+" if diff > 0 else ""
            txt_color = "#2ecc71" if diff > 0 else "#e74c3c" if diff < 0 else "#ecf0f1"
            self.diff_btn.setText(f"선택 적용 ({sign}{diff:.1f}점)")
            self.diff_btn.setStyleSheet(f"background-color: #1b263b; color: {txt_color}; font-weight: bold; border: 1px solid {txt_color};")

# ---------------------------------------------------------
# 9. 큐브 재설정 창
# ---------------------------------------------------------
class CubeDialog(QDialog):
    def __init__(self, parent, mode="윗잠"):
        super().__init__(parent)
        self.main_app = parent
        self.mode = mode
        self.setWindowTitle(f"잠재능력 재설정 ({self.mode})")
        self.setFixedSize(880, 560)

        self.before_tier = self.main_app.main_tier if mode == "윗잠" else self.main_app.add_tier
        self.before_lines = list(self.main_app.main_lines if mode == "윗잠" else self.main_app.add_lines)
        self.pity_count = self.main_app.main_pity if mode == "윗잠" else self.main_app.add_pity
        self.after_data = [None, None, None]
        self.worker = None

        self.session_rolls = 0

        self.init_ui()
        self.set_roll_view(1)
        self.update_cost_labels()
        self.update_pity_ui()
        self.update_session_stats_ui()

    def init_ui(self):
        self.root_stack = QStackedLayout()
        self.root_stack.setContentsMargins(14, 14, 14, 14)

        self.view_reroll = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.card_center_layout = QHBoxLayout()
        self.card_center_layout.setContentsMargins(0, 0, 0, 0)
        self.card_center_layout.setSpacing(14)
        self.card_center_layout.addStretch(1)

        self.before_card = PotentialCard("BEFORE", is_before=True)
        self.before_card.set_data(self.before_tier, self.before_lines, self.main_app.weapon_type == "마력", is_additional=(self.mode == "에디"))
        self.card_center_layout.addWidget(self.before_card)

        self.after_container = QWidget()
        self.after_layout = QHBoxLayout()
        self.after_layout.setContentsMargins(0, 0, 0, 0)
        self.after_layout.setSpacing(8)

        self.after_cards = []
        for i in range(3):
            c = PotentialCard("AFTER", select_callback=self.apply_option, index=i)
            self.after_cards.append(c)
            self.after_layout.addWidget(c)

        self.after_container.setLayout(self.after_layout)
        self.card_center_layout.addWidget(self.after_container)

        self.card_center_layout.addStretch(1)
        layout.addLayout(self.card_center_layout)

        self.lbl_stats = QLabel("현재 재설정 횟수: 0회 | 누적 소모 메소: 0 메소")
        self.lbl_stats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_stats.setStyleSheet("""
            background-color: #0c1017;
            color: #ffd700;
            font-size: 12px;
            font-weight: bold;
            padding: 5px;
            border-radius: 4px;
            border: 1px solid #2a374a;
        """)
        layout.addWidget(self.lbl_stats)

        pity_group = QFrame()
        pity_group.setStyleSheet("background-color: #131a24; border-radius: 6px; padding: 5px; border: 1px solid #2a374a;")
        pity_layout = QVBoxLayout()
        pity_layout.setSpacing(3)

        self.lbl_pity_text = QLabel("천장 진행도: 0 / 10회")
        self.lbl_pity_text.setStyleSheet("color: #00ffcc; font-size: 11px; font-weight: bold;")
        pity_layout.addWidget(self.lbl_pity_text)

        self.pbar_pity = QProgressBar()
        self.pbar_pity.setFixedHeight(18)
        self.pbar_pity.setTextVisible(True)
        self.pbar_pity.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pbar_pity.setStyleSheet("""
            QProgressBar {
                border: 1px solid #3e5068;
                border-radius: 4px;
                background-color: #080c12;
                color: #ffffff;
                font-size: 10px;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #16a085, stop:1 #2ecc71);
                border-radius: 3px;
            }
        """)
        pity_layout.addWidget(self.pbar_pity)
        pity_group.setLayout(pity_layout)
        layout.addWidget(pity_group)

        ctrl_group = QGroupBox("재설정 실행")
        ctrl_vbox = QVBoxLayout()
        ctrl_vbox.setSpacing(5)

        manual_box = QHBoxLayout()
        self.btn_roll_1 = QPushButton("재설정 1회")
        self.btn_roll_1.setObjectName("main_action_btn")
        self.btn_roll_1.clicked.connect(lambda: self.roll_manual(1))
        manual_box.addWidget(self.btn_roll_1)

        self.btn_roll_3 = QPushButton("재설정 3회")
        self.btn_roll_3.setObjectName("main_action_btn")
        self.btn_roll_3.clicked.connect(lambda: self.roll_manual(3))
        manual_box.addWidget(self.btn_roll_3)
        ctrl_vbox.addLayout(manual_box)

        target_box = QHBoxLayout()
        target_box.addWidget(QLabel("목표 옵션 지정:"))

        dropdown_list = MAIN_TARGET_OPTIONS if self.mode == "윗잠" else ADD_TARGET_OPTIONS
        self.cb_targets = []
        for _ in range(3):
            cb = QComboBox()
            cb.addItems(dropdown_list)
            cb.setCurrentText("선택 안 함 (무관)")
            self.cb_targets.append(cb)
            target_box.addWidget(cb)
        ctrl_vbox.addLayout(target_box)

        auto_box = QHBoxLayout()
        self.chk_stop_better = QCheckBox("현재(BEFORE)보다 점수 높으면 정지")
        self.chk_stop_better.setChecked(True)
        self.chk_stop_better.setStyleSheet("color: #f1c40f; font-weight: bold;")
        auto_box.addWidget(self.chk_stop_better)

        self.btn_auto = QPushButton("고속 자동 재설정 (3회 단위)")
        self.btn_auto.clicked.connect(self.start_auto)
        auto_box.addWidget(self.btn_auto)

        self.btn_stop = QPushButton("정지")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_auto)
        auto_box.addWidget(self.btn_stop)
        ctrl_vbox.addLayout(auto_box)

        ctrl_group.setLayout(ctrl_vbox)
        layout.addWidget(ctrl_group)

        self.view_reroll.setLayout(layout)
        self.root_stack.addWidget(self.view_reroll)

        # 단독 옵션 선택 완료 뷰
        self.view_result = QWidget()
        res_layout = QVBoxLayout()
        res_layout.setContentsMargins(20, 20, 20, 20)
        res_layout.setSpacing(15)

        lbl_res_title = QLabel("★ 잠재능력 설정이 완료되었습니다 ★")
        lbl_res_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_res_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #f1c40f;")
        res_layout.addWidget(lbl_res_title)

        res_card_box = QHBoxLayout()
        res_card_box.addStretch(1)
        self.result_single_card = PotentialCard("선택된 옵션", is_before=True)
        res_card_box.addWidget(self.result_single_card)
        res_card_box.addStretch(1)
        res_layout.addLayout(res_card_box)

        btn_res_box = QHBoxLayout()
        btn_close_dlg = QPushButton("확인 (재설정 창 닫기)")
        btn_close_dlg.setObjectName("main_action_btn")
        btn_close_dlg.clicked.connect(self.accept)
        btn_res_box.addWidget(btn_close_dlg)

        btn_continue_reroll = QPushButton("다시 계속 재설정하기")
        btn_continue_reroll.clicked.connect(self.back_to_reroll_view)
        btn_res_box.addWidget(btn_continue_reroll)
        res_layout.addLayout(btn_res_box)

        self.view_result.setLayout(res_layout)
        self.root_stack.addWidget(self.view_result)

        self.setLayout(self.root_stack)

    def set_roll_view(self, count):
        if count == 1:
            self.after_cards[0].setVisible(True)
            self.after_cards[1].setVisible(False)
            self.after_cards[2].setVisible(False)
            self.after_container.setFixedWidth(190)
        else:
            self.after_cards[0].setVisible(True)
            self.after_cards[1].setVisible(True)
            self.after_cards[2].setVisible(True)
            self.after_container.setFixedWidth(190 * 3 + 16)

    def update_cost_labels(self):
        cost = ROLL_COSTS[self.mode][self.before_tier]
        self.btn_roll_1.setText(f"재설정 1회 ({cost:,} 메소)")
        self.btn_roll_3.setText(f"재설정 3회 ({cost*3:,} 메소)")

    def update_session_stats_ui(self):
        self.lbl_stats.setText(f"현재 재설정 횟수: {self.session_rolls:,}회 | 누적 소모 메소: {self.main_app.total_spent_meso:,} 메소")

    def update_pity_ui(self):
        if self.before_tier == "레전드리":
            self.lbl_pity_text.setText("천장 상태: 최고 등급(레전드리) 도달 완료")
            self.pbar_pity.setMaximum(100)
            self.pbar_pity.setValue(100)
            self.pbar_pity.setFormat("최고 등급 달성")
            return

        limit = CEILING_LIMITS[self.mode].get(self.before_tier, 100)
        self.pbar_pity.setMaximum(limit)
        cur_val = min(self.pity_count, limit)
        self.pbar_pity.setValue(cur_val)
        pct = (cur_val / limit) * 100
        self.pbar_pity.setFormat(f"{cur_val} / {limit} ({pct:.1f}%)")
        next_tier = TIER_ORDER[TIER_ORDER.index(self.before_tier) + 1]
        self.lbl_pity_text.setText(f"{self.before_tier} → {next_tier} 등업 천장: {cur_val} / {limit}회 (달성 시 100% 확정 등업)")

    def generate_single_result(self, tier):
        self.pity_count += 1
        upgraded = False
        if tier != "레전드리":
            limit = CEILING_LIMITS[self.mode].get(tier, 999)
            if random.random() < UPGRADE_RATES.get(tier, 0) or self.pity_count >= limit:
                tier = TIER_ORDER[TIER_ORDER.index(tier) + 1]
                self.pity_count = 0
                upgraded = True

        lines = roll_three_lines(self.mode, tier)
        return tier, lines, upgraded

    def roll_manual(self, count):
        self.set_roll_view(count)
        cost_unit = ROLL_COSTS[self.mode][self.before_tier]
        self.main_app.add_spent_meso(cost_unit * count)
        self.session_rolls += count
        self.update_session_stats_ui()

        is_magic = (self.main_app.weapon_type == "마력")
        is_add = (self.mode == "에디")
        base_score = self.before_card.score
        any_upgrade = False

        for i in range(count):
            t, l, up = self.generate_single_result(self.before_tier)
            self.after_data[i] = (t, l)
            self.after_cards[i].set_data(t, l, is_magic, base_score, is_upgrade=up, is_additional=is_add)
            if up:
                any_upgrade = True

        self.update_pity_ui()

        if any_upgrade:
            self.btn_roll_1.setEnabled(False)
            self.btn_roll_3.setEnabled(False)
            self.btn_auto.setEnabled(False)

    def start_auto(self):
        if self.before_tier != "레전드리":
            QMessageBox.warning(self, "자동 재설정 제한", "자동 재설정 기능은 [레전드리] 등급에서만 작동해!\n수제로 레전드리까지 등업 후 이용해줘.")
            return

        self.set_roll_view(3)
        self.btn_auto.setEnabled(False)
        self.btn_roll_1.setEnabled(False)
        self.btn_roll_3.setEnabled(False)
        self.btn_stop.setEnabled(True)

        target_conditions = [cb.currentText() for cb in self.cb_targets]
        is_magic = (self.main_app.weapon_type == "마력")

        self.worker = AutoRollerThread(
            self.mode,
            self.before_tier,
            self.pity_count,
            target_conditions,
            self.chk_stop_better.isChecked(),
            self.before_card.score,
            is_magic
        )
        self.worker.roll_tick.connect(self.on_auto_tick)
        self.worker.finished_roll.connect(self.on_auto_finished)
        self.worker.start()

    def stop_auto(self):
        if self.worker:
            self.worker.stop()
            self.worker.wait()
        self.btn_auto.setEnabled(True)
        self.btn_roll_1.setEnabled(True)
        self.btn_roll_3.setEnabled(True)
        self.btn_stop.setEnabled(False)

    def on_auto_tick(self, data):
        self.main_app.add_spent_meso(data["cost"])
        self.session_rolls += data["rolls_count"]
        self.update_session_stats_ui()
        self.pity_count = data["pity"]
        self.update_pity_ui()

        is_magic = (self.main_app.weapon_type == "마력")
        is_add = (self.mode == "에디")
        base_score = self.before_card.score

        for i in range(3):
            cand = data["candidates"][i]
            self.after_cards[i].set_data(cand[0], cand[1], is_magic, base_score, is_additional=is_add)
            self.after_data[i] = (cand[0], cand[1])

    def on_auto_finished(self, data):
        self.stop_auto()
        is_magic = (self.main_app.weapon_type == "마력")
        is_add = (self.mode == "에디")
        base_score = self.before_card.score

        for i in range(3):
            cand = data["candidates"][i]
            self.after_cards[i].set_data(cand[0], cand[1], is_magic, base_score, is_additional=is_add)
            self.after_data[i] = (cand[0], cand[1])

        sat_idx = data.get("satisfied_idx", 0)
        self.after_cards[sat_idx].show_blind()

    def apply_option(self, index):
        if self.after_data[index]:
            tier, lines = self.after_data[index]
            self.before_tier = tier
            self.before_lines = lines
            is_add = (self.mode == "에디")
            self.before_card.set_data(tier, lines, self.main_app.weapon_type == "마력", is_additional=is_add)
            self.update_cost_labels()
            self.update_pity_ui()

            if self.mode == "윗잠":
                self.main_app.main_tier = tier
                self.main_app.main_lines = lines
                self.main_app.main_pity = self.pity_count
                self.main_app.cb_main_tier_start.setCurrentText(tier)
            else:
                self.main_app.add_tier = tier
                self.main_app.add_lines = lines
                self.main_app.add_pity = self.pity_count
                self.main_app.cb_add_tier_start.setCurrentText(tier)

            self.main_app.refresh_weapon_ui()

            self.btn_roll_1.setEnabled(True)
            self.btn_roll_3.setEnabled(True)
            self.btn_auto.setEnabled(True)

            self.result_single_card.set_data(tier, lines, self.main_app.weapon_type == "마력", is_additional=is_add)
            self.result_single_card.diff_btn.setText("적용 완료")
            self.root_stack.setCurrentIndex(1)

    def back_to_reroll_view(self):
        self.set_roll_view(1)
        self.root_stack.setCurrentIndex(0)

# ---------------------------------------------------------
# 10. 메인 윈도우 (메이플 무기 툴팁 UI)
# ---------------------------------------------------------
class MapleWeaponTooltipSimulator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("데스티니 무기 잠재능력 시뮬레이터")
        self.resize(520, 680)

        self.weapon_type = "공격력"
        self.total_spent_meso = 0

        self.main_tier = "레어"
        self.main_lines = roll_three_lines("윗잠", "레어")
        self.main_pity = 0

        self.add_tier = "레어"
        self.add_lines = roll_three_lines("에디", "레어")
        self.add_pity = 0

        self.init_ui()
        self.apply_theme()
        self.refresh_weapon_ui()

    def get_weapon_name(self):
        return "데스티니 초극검 (+10)" if self.weapon_type == "공격력" else "데스티니 카르타 (+10)"

    def add_spent_meso(self, val):
        self.total_spent_meso += val
        self.lbl_meso.setText(f"누적 소모 메소: {self.total_spent_meso:,} 메소")

    def init_ui(self):
        central_widget = QWidget()
        root_layout = QVBoxLayout()
        root_layout.setContentsMargins(12, 10, 12, 10)
        root_layout.setSpacing(6)

        top_group = QGroupBox("장비 및 초기 등급 세팅")
        top_layout = QVBoxLayout()
        top_layout.setContentsMargins(8, 6, 8, 6)
        top_layout.setSpacing(4)

        row1 = QHBoxLayout()
        self.cb_type = QComboBox()
        self.cb_type.addItems(["공격력 무기 (데스티니 초극검)", "마력 무기 (데스티니 카르타)"])
        self.cb_type.currentIndexChanged.connect(self.on_type_change)
        row1.addWidget(self.cb_type, 3)

        btn_save = QPushButton("무기 저장")
        btn_save.clicked.connect(self.save_json)
        row1.addWidget(btn_save, 1)

        btn_load = QPushButton("무기 로드")
        btn_load.clicked.connect(self.load_json)
        row1.addWidget(btn_load, 1)
        top_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("윗잠 등급:"))
        self.cb_main_tier_start = QComboBox()
        self.cb_main_tier_start.addItems(TIER_ORDER)
        self.cb_main_tier_start.setCurrentText("레어")
        self.cb_main_tier_start.currentTextChanged.connect(self.on_main_tier_manual_change)
        row2.addWidget(self.cb_main_tier_start)

        row2.addWidget(QLabel("에디 등급:"))
        self.cb_add_tier_start = QComboBox()
        self.cb_add_tier_start.addItems(TIER_ORDER)
        self.cb_add_tier_start.setCurrentText("레어")
        self.cb_add_tier_start.currentTextChanged.connect(self.on_add_tier_manual_change)
        row2.addWidget(self.cb_add_tier_start)
        top_layout.addLayout(row2)

        top_group.setLayout(top_layout)
        root_layout.addWidget(top_group)

        self.item_tooltip = QFrame()
        self.item_tooltip.setObjectName("item_tooltip")
        tt_layout = QVBoxLayout()
        tt_layout.setContentsMargins(12, 10, 12, 10)
        tt_layout.setSpacing(4)

        self.lbl_starforce = QLabel("★ ★ ★ ★ ★  ★ ★ ★ ★ ★  ★ ★ ★ ★ ★  ★ ★ ★ ★ ★  ★ ★ ★ ★ ★")
        self.lbl_starforce.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_starforce.setStyleSheet("color: #ffcc00; font-size: 10px;")
        tt_layout.addWidget(self.lbl_starforce)

        self.lbl_name = QLabel(self.get_weapon_name())
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_name.setStyleSheet("color: #ff9900; font-size: 16px; font-weight: bold;")
        tt_layout.addWidget(self.lbl_name)

        self.lbl_sub_grade = QLabel("(레어 아이템)")
        self.lbl_sub_grade.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sub_grade.setStyleSheet("color: #3498db; font-size: 11px;")
        tt_layout.addWidget(self.lbl_sub_grade)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.HLine)
        sep1.setStyleSheet("color: #4a5568;")
        tt_layout.addWidget(sep1)

        self.lbl_specs = QLabel(
            "• 무기분류: 한손검   • 공격속도: 빠름 (4등급)\n"
            "• STR: +160 (+40 +120)   • DEX: +160 (+40 +120)\n"
            "• 공격력: +754 (+285 +469)\n"
            "• 보스 몬스터 공격 시 데미지 +30%   • 몬스터 방어율 무시 +20%"
        )
        self.lbl_specs.setStyleSheet("color: #ecf0f1; font-size: 11px; line-height: 1.3;")
        tt_layout.addWidget(self.lbl_specs)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet("color: #4a5568;")
        tt_layout.addWidget(sep2)

        self.lbl_main_pot_title = QLabel("잠재옵션")
        self.lbl_main_pot_title.setStyleSheet("font-size: 12px; font-weight: bold;")
        tt_layout.addWidget(self.lbl_main_pot_title)

        self.main_line_labels = []
        for _ in range(3):
            lbl = QLabel("-")
            self.main_line_labels.append(lbl)
            tt_layout.addWidget(lbl)

        sep3 = QFrame()
        sep3.setFrameShape(QFrame.Shape.HLine)
        sep3.setStyleSheet("color: #4a5568;")
        tt_layout.addWidget(sep3)

        self.lbl_add_pot_title = QLabel("에디셔널 잠재옵션")
        self.lbl_add_pot_title.setStyleSheet("font-size: 12px; font-weight: bold;")
        tt_layout.addWidget(self.lbl_add_pot_title)

        self.add_line_labels = []
        for _ in range(3):
            lbl = QLabel("-")
            self.add_line_labels.append(lbl)
            tt_layout.addWidget(lbl)

        self.item_tooltip.setLayout(tt_layout)
        root_layout.addWidget(self.item_tooltip)

        self.lbl_meso = QLabel("누적 소모 메소: 0 메소")
        self.lbl_meso.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_meso.setStyleSheet("font-size: 13px; font-weight: bold; color: #ffd700; margin: 2px 0;")
        root_layout.addWidget(self.lbl_meso)

        btn_box = QHBoxLayout()
        btn_box.setSpacing(6)
        btn_open_main = QPushButton("윗잠재능력 재설정 창 열기")
        btn_open_main.setObjectName("main_action_btn")
        btn_open_main.clicked.connect(lambda: self.open_cube_dialog("윗잠"))
        btn_box.addWidget(btn_open_main)

        btn_open_add = QPushButton("에디셔널 재설정 창 열기")
        btn_open_add.setObjectName("main_action_btn")
        btn_open_add.clicked.connect(lambda: self.open_cube_dialog("에디"))
        btn_box.addWidget(btn_open_add)
        root_layout.addLayout(btn_box)

        central_widget.setLayout(root_layout)
        self.setCentralWidget(central_widget)

    def on_type_change(self, idx):
        self.weapon_type = "공격력" if idx == 0 else "마력"
        self.lbl_name.setText(self.get_weapon_name())
        if self.weapon_type == "마력":
            self.lbl_specs.setText(
                "• 무기분류: 스태프   • 공격속도: 보통 (6등급)\n"
                "• INT: +160 (+40 +120)   • LUK: +160 (+40 +120)\n"
                "• 마력: +754 (+285 +469)\n"
                "• 보스 몬스터 공격 시 데미지 +30%   • 몬스터 방어율 무시 +20%"
            )
        else:
            self.lbl_specs.setText(
                "• 무기분류: 한손검   • 공격속도: 빠름 (4등급)\n"
                "• STR: +160 (+40 +120)   • DEX: +160 (+40 +120)\n"
                "• 공격력: +754 (+285 +469)\n"
                "• 보스 몬스터 공격 시 데미지 +30%   • 몬스터 방어율 무시 +20%"
            )
        self.refresh_weapon_ui()

    def on_main_tier_manual_change(self, tier):
        self.main_tier = tier
        self.main_pity = 0
        self.main_lines = roll_three_lines("윗잠", tier)
        self.refresh_weapon_ui()

    def on_add_tier_manual_change(self, tier):
        self.add_tier = tier
        self.add_pity = 0
        self.add_lines = roll_three_lines("에디", tier)
        self.refresh_weapon_ui()

    def refresh_weapon_ui(self):
        m_color = TIER_COLORS.get(self.main_tier, "#3498db")
        self.lbl_main_pot_title.setText(f"잠재옵션 ({self.main_tier})")
        self.lbl_main_pot_title.setStyleSheet(f"color: {m_color}; font-size: 12px; font-weight: bold;")
        self.lbl_sub_grade.setText(f"({self.main_tier} 아이템)")
        self.lbl_sub_grade.setStyleSheet(f"color: {m_color}; font-size: 11px;")

        for i in range(3):
            text, badge = self.main_lines[i]
            line_color = TIER_COLORS.get("레전드리" if badge == "L" else "유니크" if badge == "U" else "에픽" if badge == "E" else "레어", "#3498db")
            self.main_line_labels[i].setText(f"+ {text}")
            self.main_line_labels[i].setStyleSheet(f"color: {line_color}; font-size: 11px;")

        a_color = TIER_COLORS.get(self.add_tier, "#3498db")
        self.lbl_add_pot_title.setText(f"에디셔널 잠재옵션 ({self.add_tier})")
        self.lbl_add_pot_title.setStyleSheet(f"color: {a_color}; font-size: 12px; font-weight: bold;")

        for i in range(3):
            text, badge = self.add_lines[i]
            line_color = TIER_COLORS.get("레전드리" if badge == "L" else "유니크" if badge == "U" else "에픽" if badge == "E" else "레어", "#3498db")
            self.add_line_labels[i].setText(f"+ {text}")
            self.add_line_labels[i].setStyleSheet(f"color: {line_color}; font-size: 11px;")

    def open_cube_dialog(self, mode):
        dlg = CubeDialog(self, mode)
        dlg.exec()

    def save_json(self):
        path, _ = QFileDialog.getSaveFileName(self, "무기 저장", "", "JSON Files (*.json)")
        if path:
            data = {
                "weapon_type": self.weapon_type,
                "main_tier": self.main_tier,
                "main_lines": self.main_lines,
                "main_pity": self.main_pity,
                "add_tier": self.add_tier,
                "add_lines": self.add_lines,
                "add_pity": self.add_pity,
                "total_spent_meso": self.total_spent_meso
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)

    def load_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "무기 불러오기", "", "JSON Files (*.json)")
        if path:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.weapon_type = data.get("weapon_type", "공격력")
            self.cb_type.setCurrentIndex(0 if self.weapon_type == "공격력" else 1)
            self.main_tier = data.get("main_tier", "레어")
            self.main_lines = data.get("main_lines", [])
            self.main_pity = data.get("main_pity", 0)
            self.add_tier = data.get("add_tier", "레어")
            self.add_lines = data.get("add_lines", [])
            self.add_pity = data.get("add_pity", 0)
            self.total_spent_meso = data.get("total_spent_meso", 0)

            self.cb_main_tier_start.blockSignals(True)
            self.cb_add_tier_start.blockSignals(True)
            self.cb_main_tier_start.setCurrentText(self.main_tier)
            self.cb_add_tier_start.setCurrentText(self.add_tier)
            self.cb_main_tier_start.blockSignals(False)
            self.cb_add_tier_start.blockSignals(False)

            self.lbl_meso.setText(f"누적 소모 메소: {self.total_spent_meso:,} 메소")
            self.refresh_weapon_ui()

    def apply_theme(self):
        self.setStyleSheet("""
            * {
                font-family: 'SUIT Light', 'SUIT', sans-serif;
            }
            QMainWindow, QDialog {
                background-color: #0f141c;
            }
            #item_tooltip {
                background-color: rgba(10, 14, 22, 0.95);
                border: 2px solid #b38f4d;
                border-radius: 8px;
            }
            #card_frame {
                background-color: #1a2332;
                border: 1px solid #364963;
                border-radius: 6px;
            }
            #diff_button {
                background-color: #24334a;
                border-radius: 4px;
                padding: 4px;
                font-size: 11px;
            }
            #diff_button:hover {
                border: 1px solid #f39c12;
            }
            #blind_single_btn {
                background: qradialgradient(cx: 0.5, cy: 0.5, radius: 0.8, fx: 0.5, fy: 0.5, stop: 0 #1b629b, stop: 1 #0a1c2e);
                border: 2px solid #3498db;
                border-radius: 6px;
                color: #ffffff;
                font-size: 58px;
                font-weight: bold;
                padding: 0px;
                margin: 0px;
            }
            #blind_single_btn:hover {
                border: 2px solid #5dade2;
            }
            #main_action_btn {
                background-color: #1f618d;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 7px;
                font-size: 12px;
            }
            #main_action_btn:hover {
                background-color: #2980b9;
            }
            #main_action_btn:disabled {
                background-color: #3e4856;
                color: #7f8c8d;
            }
            QGroupBox {
                color: #f39c12;
                font-weight: bold;
                border: 1px solid #2f3b4c;
                margin-top: 4px;
                padding-top: 4px;
                border-radius: 4px;
                font-size: 11px;
            }
            QLabel {
                color: #ffffff;
            }
            QPushButton {
                background-color: #2c3e50;
                color: white;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #3e5871;
            }
            QPushButton:disabled {
                background-color: #1a222d;
                color: #555555;
            }
            QComboBox {
                background-color: #1e293b;
                color: white;
                border: 1px solid #3e4856;
                padding: 2px 4px;
                border-radius: 3px;
                font-size: 11px;
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MapleWeaponTooltipSimulator()
    window.show()
    sys.exit(app.exec())