from dataclasses import dataclass

@dataclass
class FormulaData:
    """一個用來儲存單一儲存格公式資訊的資料類別。"""
    address: str
    formula: str
    value: any = None # 可選的儲存格值
