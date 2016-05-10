namespace BJExcelLib.Finance

open BJExcelLib.Math

/// Yield curve class that can implement a custom interpolation function on (transforms of) discount factors. To do so, the dfInterpolator function
/// has to be provided. This higher-order function takes as input an array of ttm and yield tuples and returns a function that must associate with
/// a time to maturity a discount factor. Note that internally theInterpolator can implement transformations of yields, but then it must also provide
/// inverse transform on the interpolated quantity.
type YieldCurveDFGeneral(definingPoints:seq<YieldCurvePoint>, referenceDate, dayCounter, ?compoundingFrequency,?dfInterpolator ) as this =
    inherit YieldCurve(definingPoints, referenceDate, dayCounter, (defaultArg compoundingFrequency Continuous))

    let xydata =
        (this :> IYieldCurve).DefiningPoints
        |> Seq.map (fun ycp -> ycp.YearFraction,(YieldCurveHelper.YieldToDiscountFactor ycp.Yield ycp.YearFraction (this :> IYieldCurve).CompoundingFrequency))
        |> Array.ofSeq

    /// create interpolating function
    let defaultInterpolator = Interpolation.interpolatorPiecewiseLinear(xydata)
    do if (Option.isNone defaultInterpolator) && (Array.length xydata > 1) then
        raise(new System.ArithmeticException("Could not create interpolation function for data"))
    let interpolator = defaultArg dfInterpolator (Option.get defaultInterpolator)

    /// Internal function to determine discount factors is in this case derived from the internal yield calculation           
    override __.InternalDiscountFactorCalculation(validatedTime) = interpolator validatedTime