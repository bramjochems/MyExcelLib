namespace BJExcelLib.Finance

open System
open BJExcelLib.Math
open BJExcelLib.Util

module public YieldCurveHelper =
    let validateTTM t = t >= 0.

    /// Converts a yield for a given time to maturity and given compounding frequency to a discount factor
    let YieldToDiscountFactor (yield_:float) timeToMaturity compoundingFrequency =
        match compoundingFrequency with
        | TimesPerYear(n) when n < 1 -> raise(ArgumentOutOfRangeException("Invalid compounding frequency argument in YieldToDiscountFactor"))
        | Continuous -> Math.Exp(yield_ * timeToMaturity)
        | TimesPerYear(n)->
            let n = (float) n
            Math.Pow(1. + yield_/n,-n*timeToMaturity)

    /// Converts a discount factor for a given time to maturity and compounding frequency to a yield that uses the same compounding frequency
    let DiscountFactorToYield discountFactor timeToMaturity compoundingFrequency =
        if timeToMaturity < 0. then raise(ArgumentOutOfRangeException("Cannot determine yield from discount factor for time < 0"))
        if discountFactor <= 0. then raise(ArgumentOutOfRangeException("Cannot deal with non-positive discount factor")) 
        if timeToMaturity = 0. then 0.
        else match compoundingFrequency with
             | TimesPerYear(n) when n < 1 -> raise(ArgumentOutOfRangeException("Invalid compounding frequency argument in YieldToDiscountFactor"))
             | Continuous -> -Math.Log(discountFactor) / timeToMaturity
             | TimesPerYear(n) ->
                let n = (float n)
                (Math.Pow(discountFactor, 1./(n*timeToMaturity))-1.)*n


/// Abstract base class for yield curves
[<AbstractClass>]
type YieldCurve(definingPoints:seq<YieldCurvePoint>, referenceDate, dayCounter:DateTime->float, ?compoundingFrequency ) =

    /// preprocess the defining points
    let definingPoints =
        definingPoints
         |> Seq.sortBy (fun ycp -> ycp.YearFraction)
         |> Seq.distinctBy( fun ycp -> ycp.YearFraction)
         |> Seq.toList

    do if definingPoints = [] then raise(new ArgumentException("To few points to define curve"))

    /// Abstract member that subclasses must implement
    abstract member InternalDiscountFactorCalculation : float -> float

    /// Duplicate definition of yield function in class rather than interface.
    /// This is is a bit cluncky but the way to deal with virtual methods
    abstract Yield : timeToMaturity:float -> float

    /// Default implementation of Yield method
    default __.Yield(timeToMaturity) =
        let iyc = (__ :> IYieldCurve)
        YieldCurveHelper.DiscountFactorToYield (iyc.Discount(timeToMaturity)) timeToMaturity iyc.CompoundingFrequency

    /// Implementation of the IYieldcurve interface
    interface IYieldCurve with
        /// Getter to the daycounter function
        member __.Daycounter
            with get() = dayCounter

        /// Getter to the reference date for the yield curve
        member __.ReferenceDate
            with get() = referenceDate

        /// Getter to a boolean that indicates whether the curve
        member __.CompoundingFrequency
            with get() = defaultArg compoundingFrequency Continuous

        /// Gets the points that define the curve
        member __.DefiningPoints
            with get() = definingPoints

        /// Return the discount factor for a given timeToMaturity
        member __.Discount(timeToMaturity) =
            match YieldCurveHelper.validateTTM timeToMaturity with
            | true -> __.InternalDiscountFactorCalculation timeToMaturity
            | false -> 1.

        /// Yield for time to maturity
        member __.Yield(timeToMaturity) =
            __.Yield(timeToMaturity) // pass to class member to enable creation of a virtual method

        /// Apply a parallel shock to a IYieldCurve
        member __.ParallelShock isAbsoluteShock shock =
            let definingPointMapper =
                match (isAbsoluteShock, shock) with
                /// first two are when a shock isn't a real shock. Still construct a shockcurve object to avoid returning a reference to this object
                | true, x when abs(x) < 1e-8     -> fun ycp -> { ycp with Yield = 0. }
                | false, x when abs(x)-1. < 1e-8 -> fun ycp -> { ycp with Yield = 0. }
                | true, x -> fun ycp -> {ycp with Yield = x}
                | false, x ->fun ycp -> {ycp with Yield = (x-1.)*ycp.Yield}
            let iyc = (__ :> IYieldCurve)
            let defPoints = iyc.DefiningPoints |> List.map definingPointMapper
            let shockCurve = YieldCurveLinearSpot(defPoints,iyc.ReferenceDate,iyc.Daycounter, iyc.CompoundingFrequency) :> IYieldCurve
            ShockedYieldCurve(iyc,shockCurve) :> IYieldCurve
        
        /// Apply a custom shock to a yieldcurve
        member __.Shock(shockCurve) =
            new ShockedYieldCurve(__:>IYieldCurve,shockCurve) :> IYieldCurve

/// Class for shocked yield curves which are sum objects of two yield curves
and ShockedYieldCurve(baseCurve,shockCurve) as this =
    inherit YieldCurve(Seq.empty, baseCurve.ReferenceDate, baseCurve.Daycounter, baseCurve.CompoundingFrequency)

    do if isNull baseCurve then raise(NullReferenceException("Null base curve provided to ShockedYieldCurve"))
       if isNull shockCurve then raise (NullReferenceException("Null shocked curve provided to ShockedYieldCurve"))
    let _baseCurve = baseCurve
    let _shockCurve = shockCurve

    /// Implement the calculation for a discount factor for shocked curve. All other functionality follows from the base class YieldCurve
    override __.InternalDiscountFactorCalculation(validatedTime) =
        _baseCurve.Discount(validatedTime)*_shockCurve.Discount(validatedTime)

/// Yield curve that interpolates linear in spot rates
and YieldCurveLinearSpot(definingPoints:seq<YieldCurvePoint>, referenceDate, dayCounter, ?compoundingFrequency ) as this =
    inherit YieldCurve(definingPoints,referenceDate,dayCounter, (defaultArg compoundingFrequency Continuous))
    /// process hte input points into an array of tuples
    let xydata =
        (this :> IYieldCurve).DefiningPoints
        |> List.map (fun dp -> dp.YearFraction,dp.Yield)
        |> Array.ofList

    /// create interpolating function
    let interpolator = Interpolation.interpolatorPiecewiseLinear(xydata)
    do if (Option.isNone interpolator) && (Array.length xydata > 1) then
        raise(new System.ArithmeticException("Could not create interpolation function for data"))
    let interpolator = Option.get interpolator
        
    /// Internal function to determine discount factors is in this case derived from the internal yield calculation           
    override __.InternalDiscountFactorCalculation(validatedTime) = 
        YieldCurveHelper.YieldToDiscountFactor (this.Yield(validatedTime)) validatedTime (this :> IYieldCurve).CompoundingFrequency

    /// Override for the yield calculation to be piecewise linear in spot
    override __.Yield(t) =
        match t with
        | x when x <= fst xydata.[0] -> snd xydata.[0] //flat extrapolation before first point
        | x when x >= fst xydata.[xydata.Length-1] -> snd xydata.[xydata.Length-1] // flat extrapolation after last point
        | _ -> interpolator(t)


