namespace BJExcelLib.Finance

/// Input points from which yield curves are construed.
type YieldCurvePoint =
    { YearFraction : float;
      Yield : float}