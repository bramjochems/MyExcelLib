namespace BJExcelLib.Finance

type CompoundingFrequency =
    | Continuous
    | TimesPerYear of int

/// Interface for yield curve objects
type IYieldCurve =
    abstract member Daycounter : (System.DateTime -> float) with get
    abstract member ReferenceDate : System.DateTime with get
    abstract member CompoundingFrequency: CompoundingFrequency with get
    abstract member DefiningPoints : YieldCurvePoint list with get
    abstract member Discount : timeToMaturity:float -> float
    abstract member Yield : timeToMaturity:float -> float
    abstract member ParallelShock : isAbsoluteShock: bool -> shock:float -> IYieldCurve
    abstract member Shock : shockCurve:IYieldCurve -> IYieldCurve