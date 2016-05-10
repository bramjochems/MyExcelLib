namespace BJExcelLib.Finance

open BJExcelLib.Math

/// Yield curve class that can implement a custom interpolation function on (transforms of) spot rates. To do so, the yieldInterpolator function
/// has to be provided. This higher-order function takes as input an array of ttm and yield tuples and returns a function that must associate with
/// a time to maturity a yield. Note that internally yieldInterpolator can implement transformations of yields, but then it must also provide the
/// inverse transform on the interpolated quantity.
type YieldCurveSpotGeneral(definingPoints:seq<YieldCurvePoint>, referenceDate, dayCounter, ?compoundingFrequency,?yieldInterpolator ) as this =
    inherit YieldCurve(definingPoints, referenceDate, dayCounter, (defaultArg compoundingFrequency Continuous))

    let xydata =
        (this :> IYieldCurve).DefiningPoints
        |> Seq.map (fun ycp -> ycp.YearFraction,(YieldCurveHelper.YieldToDiscountFactor ycp.Yield ycp.YearFraction (this :> IYieldCurve).CompoundingFrequency))
        |> Array.ofSeq

    /// create interpolating function
    let defaultInterpolator = Interpolation.interpolatorPiecewiseLinear(xydata)
    do if (Option.isNone defaultInterpolator) && (Array.length xydata > 1) then
        raise(new System.ArithmeticException("Could not create interpolation function for data"))
    let interpolator = defaultArg yieldInterpolator (Option.get defaultInterpolator)

    /// Internal function to determine discount factors is in this case derived from the internal yield calculation           
    override __.InternalDiscountFactorCalculation(validatedTime) =
        YieldCurveHelper.YieldToDiscountFactor (this.Yield(validatedTime)) validatedTime (this :> IYieldCurve).CompoundingFrequency

    /// Override for the yield calculation to be piecewise linear in spot
    override __.Yield(t) = interpolator t