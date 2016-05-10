namespace BJExcelLib.Finance

open BJExcelLib.Math

/// Yield curve class that interpolates linear on discount factor
type YieldCurveLinearDiscount(definingPoints:seq<YieldCurvePoint>, referenceDate, dayCounter, ?compoundingFrequency ) as this =
    inherit YieldCurve(definingPoints, referenceDate, dayCounter, (defaultArg compoundingFrequency Continuous))

    let xydata =
        (this :> IYieldCurve).DefiningPoints
        |> Seq.map (fun ycp -> ycp.YearFraction,(YieldCurveHelper.YieldToDiscountFactor ycp.Yield ycp.YearFraction (this :> IYieldCurve).CompoundingFrequency))
        |> Array.ofSeq

    /// create interpolating function
    let interpolator = Interpolation.interpolatorPiecewiseLinear(xydata)
    do if (Option.isNone interpolator) && (Array.length xydata > 1) then
        raise(new System.ArithmeticException("Could not create interpolation function for data"))
    let interpolator = Option.get interpolator

        /// Internal function to determine discount factors is in this case derived from the internal yield calculation           
    override __.InternalDiscountFactorCalculation(validatedTime) =
        interpolator validatedTime