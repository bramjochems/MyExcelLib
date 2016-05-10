namespace BJExcelLib.UDF

/// Contains miscelleaneous UDFs for which I haven't bothered to create modules implementing their functionality.
/// As this library expands and more modules start to exists, consider moving functions from here to a UDF module
/// that contains all related functionality.
module public Misc =
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.Util.Extensions

    let private fwdfunc (t1 : float) y1 t2 y2 =
        (y2*t2-y1*t1)/(t2-t1)

    /// Computes a year fraction between two dates on an Act/Act (ISDA) daycount
    [<ExcelFunction(Name="ForwardVolatility",Description="Calculates the forward volatility from two data points", Category="BJExcelLib.Volatility",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public ForwardVol  ([<ExcelArgument(Name="ttm1",Description = "First time to maturity")>] t1 : obj,
                            [<ExcelArgument(Name="vol1",Description = "Associated first volatility")>] v1 : obj,
                            [<ExcelArgument(Name="ttm2",Description = "Second time to maturity")>] t2 : obj,
                            [<ExcelArgument(Name="vol1",Description = "Associated second volatility")>] v2 : obj,
                            [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =            
        let ttm1, vol1, ttm2, vol2 = t1 |> validateFloat, v1 |> validateFloat, t2 |> validateFloat, v2 |> validateFloat
        if (ttm1.IsSome) && vol1.IsSome && ttm2.IsSome && vol2.IsSome && vol1.Value >= 0. && vol2.Value >= 0. && ttm1.Value >= 0. && ttm2.Value >= 0. then
            let ttm1, vol1, ttm2, vol2 = ttm1.Value, vol1.Value, ttm2.Value, vol2.Value
            let fwdvar = if ttm1 < ttm2 then fwdfunc ttm1 (vol1*vol1) ttm2 (vol2*vol2) else fwdfunc ttm2 (vol2*vol2) ttm1 (vol1*vol1)
            if fwdvar < 0. then None else Some(sqrt(fwdvar))
        else None
        |> valueToExcel

    /// Computes a year fraction between two dates on an Act/Act (ISDA) daycount
    [<ExcelFunction(Name="ForwardRate",Description="Calculates the forward rate from two data points", Category="BJExcelLib.Volatility",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public ForwardRate( [<ExcelArgument(Name="ttm1",Description = "First time to maturity")>] t1 : obj,
                            [<ExcelArgument(Name="vol1",Description = "Associated first rate")>] v1 : obj,
                            [<ExcelArgument(Name="ttm2",Description = "Second time to maturity")>] t2 : obj,
                            [<ExcelArgument(Name="vol1",Description = "Associated second rate")>] v2 : obj,
                            [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =      
        let ttm1, rate1, ttm2, rate2 = t1 |> validateFloat, v1 |> validateFloat, t2 |> validateFloat, v2 |> validateFloat
        if (ttm1.IsSome) && rate1.IsSome && ttm2.IsSome && rate2.IsSome && rate1.Value >= 0. && rate2.Value >= 0. && ttm1.Value >= 0. && ttm2.Value >= 0. then
            let ttm1, rate1, ttm2, rate2 = ttm1.Value, rate1.Value, ttm2.Value, rate2.Value
            let fwdrate = if ttm1 < ttm2 then fwdfunc ttm1 (rate1) ttm2 (rate2) else fwdfunc ttm2 (rate2) ttm1 (rate1)
            Some(fwdrate)
        else None
        |> valueToExcel

    [<ExcelFunction(Name="DiscountFactor",Description="Calculates the discount rate from a given time to maturity, discount factor and compunding frequency",Category="BJExcelLib.Misc",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public DiscountFactor(  [<ExcelArgument(Name="ttm",Description ="Time to maturity")>] t : obj,
                                [<ExcelArgument(Name="Rate",Description ="Interest reate to use")>] r : obj,
                                [<ExcelArgument(Name="Frequency",Description="The number of interest payments per year. Optional, default = 0, which is interpreted as continuous compounding")>] f : obj) =

        let ttm, rate, freq = t |> validateFloat, r |> validateFloat, f |> validateFloat |> Option.getOrElse 0. |> max 0.
        if ttm.IsSome && rate.IsSome then
            if freq = 0. then exp(-rate.Value * ttm.Value) |> Some
            else (1. + rate.Value/freq)**(-freq*ttm.Value) |> Some
        else None
        |> valueToExcel