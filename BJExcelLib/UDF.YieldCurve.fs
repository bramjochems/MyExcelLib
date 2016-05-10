namespace BJExcelLib.UDF

/// Contains UDFs for interpolation functions. Two types of functions exists. One are UDFS whose excel name is of the form Interpolate.XXX. Those
/// UDFs are UDFs that directly do interpolation. For efficiency reasons, when interpolating multiple data points, array formulae on the new x-values
/// can be used. The other are UDFs of the form BuildInterpolator.XXX. These UDFs create a function in memory which can later be used for interpolation
/// by calling Interpolate.FromObject. The idea here is that array formulae are never needed. These functions are useful if the input data is relatively
/// fixed but the value to interpolate at can change one-by-one (in which case recalculating all of them is inefficient).
module public YieldCurves =
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.ExcelDna.Cache
    open BJExcelLib.Finance
    
    let private YC_TAG = "BJYieldCurve"
    let private validateArray = array1DAsArray validateFloat

    // Registers an interest rate curve in teh cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="RegisterYieldCurve.LinearSpot", Description="Registers an yieldcurve object that interpolates piecewise linear on spot rates",Category="BJExcelLib.YieldCurves")>]
    let public udfregyclinspot  ([<ExcelArgument(Name="ttm",Description = "Time to maturities")>] ttm : obj [],
                                 [<ExcelArgument(Name="points",Description = "spot rates points")>] yields_ : obj []) =
        let ttm = validateArray ttm
        let yields_ = validateArray yields_
        
        if Array.length ttm = Array.length yields_ then
            let dp =
                Array.zip ttm yields_
                |> Seq.filter (fun (t,y) -> Option.isSome t && Option.isSome y)
                |> Seq.map (fun (t,y) -> Option.get t,Option.get y)
                |> Seq.filter (fun (t,_) -> t>= 0.)
                |> Seq.map (fun (ttm,y) -> {YearFraction = ttm; Yield = y})
            let today = System.DateTime.Today
            YieldCurveLinearSpot(dp,today, (fun date -> Date.yearfrac Actual_365qrt today date), Continuous) |> Some
        else None
        |> fun x -> match x with
                    |None -> None |> valueToExcel
                    |Some(yc) -> register YC_TAG yc
        

    // Registers an interest rate curve in teh cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="RegisterYieldCurve.LinearDiscountFactor", Description="Registers an yieldcurve object that interpolates piecewise linear on discount factors",Category="BJExcelLib.YieldCurves")>]
    let public udfregyclindf  ([<ExcelArgument(Name="ttm",Description = "Time to maturities")>] ttm : obj [],
                               [<ExcelArgument(Name="points",Description = "spot rates points")>] yields_ : obj []) =
        let ttm = validateArray ttm
        let yields_ = validateArray yields_
        
        if Array.length ttm = Array.length yields_ then
            let dp =
                Array.zip ttm yields_
                |> Seq.filter (fun (t,y) -> Option.isSome t && Option.isSome y)
                |> Seq.map (fun (t,y) -> Option.get t,Option.get y)
                |> Seq.filter (fun (t,_) -> t>= 0.)
                |> Seq.map (fun (ttm,y) -> {YearFraction = ttm; Yield = y})
            let today = System.DateTime.Today
            YieldCurveLinearDiscount(dp,today, (fun date -> Date.yearfrac Actual_365qrt today date), Continuous) |> Some
        else None
        |> fun x -> match x with
                    |None -> None |> valueToExcel
                    |Some(yc) -> register YC_TAG yc
        

    /// Looks up the yield for a yieldcurve at a given point
    /// in the cache is that you avoid array formulas and trigger only a single calculation if a single value changes.
    [<ExcelFunction(Name="YieldCurve.Yield", Description="Retrieves yield from previously registered yield curve object",Category="BJExcelLib.YieldCurve")>]
    let public udfycyfromobj ([<ExcelArgument(Name="YieldCurve",Description="Name of the yield curve")>] name : obj,
                              [<ExcelArgument(Name="Tnew",Description = "New time value to get the yield for")>] xnew : obj) =
        validateString name
            |> fun s -> if s.IsSome && (s.Value.Contains(YC_TAG)) then (Option.bind lookup s) else None
            |> fun s -> match s with
                        | None      -> None
                        | Some(f)   -> xnew |> validateFloat |> Option.map (f :?> IYieldCurve).Yield
            |> valueToExcel     

    /// Looks up the discount factor for 
    /// in the cache is that you avoid array formulas and trigger only a single calculation if a single value changes.
    [<ExcelFunction(Name="YieldCurve.DiscountFactor", Description="Retrieves the DiscountFactor from previously registered yield curve object",Category="BJExcelLib.YieldCurve")>]
    let public udfycdffromobj ([<ExcelArgument(Name="YieldCurve",Description="Name of the yield curve")>] name : obj,
                               [<ExcelArgument(Name="Tnew",Description = "New time value to get the yield for")>] xnew : obj) =
        validateString name
            |> fun s -> if s.IsSome && (s.Value.Contains(YC_TAG)) then (Option.bind lookup s) else None
            |> fun s -> match s with
                        | None      -> None
                        | Some(f)   -> xnew |> validateFloat |> Option.map (f :?> IYieldCurve).DiscountFactor
            |> valueToExcel 