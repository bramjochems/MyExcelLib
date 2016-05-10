namespace BJExcelLib.Finance

open System

[<AutoOpen>]
module YieldCurveExtensions =
    
    type IYieldCurve with

        /// Returns a discount factor for a date rater than a time to maturity
        member __.Discount(date:DateTime) =
            __.DiscountFactor(__.Daycounter(date))

        /// Returns a yield for a date rather than a time to maturity
        member __.Yield(date:DateTime) =
            __.Yield(__.Daycounter(date))

        /// Returns a discount factor for a time to maturity for a yieldcurve where a spread gets applied to the curve
        member __.Discount(ttm,spread) =
            match __.CompoundingFrequency with
            | Continuous -> __.DiscountFactor(ttm)* (YieldCurveHelper.YieldToDiscountFactor spread ttm Continuous)
            | _ as x -> YieldCurveHelper.YieldToDiscountFactor (__.Yield(ttm) + spread) ttm x

        /// Calculates the forward rate between two times to maturities
        member __.ForwardRate(t1,t2) =
            match abs(t1-t2) > 1e-8, __.CompoundingFrequency with
            | false, _ -> 0.
            | true, Continuous -> (__.Yield(t2) - __.Yield(t1))/(t2-t1)
            | true, _ as x -> Math.Pow(Math.Pow(1. + __.Yield(t2), t2) / Math.Pow(1. + __.Yield(t1), t1), 1. / (t2 - t1)) - 1.;