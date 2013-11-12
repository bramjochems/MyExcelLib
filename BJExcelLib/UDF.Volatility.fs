namespace BJFinanceLib.Excel.UDF

/// Contains excel UDFs for various volatility related functions
module public Volatility =
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.Finance

    let private INTERPOLATOR_TAG = "VolInterpolator"

    let private SVI_CALIB_CHECK_FOR_ARB = true
    let private SVI_CALIB_MIN_SIGMA = 0.001

    let private validate5tuple (a,b,c,d,e) =
        (a |> validateFloat, b |> validateFloat, c |> validateFloat, d |> validateFloat, e|> validateFloat)
            |> fun (v,w,x,y,z) -> if v.IsNone || w.IsNone || x.IsNone || y.IsNone || z.IsNone then None
                                  else Some(v.Value,w.Value,x.Value,y.Value,z.Value)
    let private validaterawparams a b m r s =
        validate5tuple(a,b,m,r,s)
    
    /// Validates jump-wing parameters. En passent converts atm vol and min vol to variance as well
    let private validatejwparameters atm skew put call minv =
        validate5tuple(atm,skew,put,call,minv)
            |> Option.bind (fun (atm,skew,put,call,minv) -> if atm <= 0. || minv <= 0. then None
                                                            else Some(atm*atm,skew,put,call,minv*minv))

    [<ExcelFunction(Name="SVI.RawVolatility",Description="Volatility for given strike from SVI Raw parameters",Category="BJFinanceLib.Volatility",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDF_SVI_RAW_Vol([<ExcelArgument(Name="a",Description="The a parameter in the SVI Raw parametrization of total variance")>] a : obj,
                        [<ExcelArgument(Name="b",Description="The b parameter in the SVI Raw parametrization of total variance")>] b : obj,
                        [<ExcelArgument(Name="r",Description="The rho parameter in the SVI Raw parametrization of total variance")>] r : obj,
                        [<ExcelArgument(Name="m",Description="The mu parameter in the SVI Raw parametrization of total variance")>] m : obj,
                        [<ExcelArgument(Name="s",Description="The sigma parameter in the SVI Raw parametrization of total variance")>] s : obj,
                        [<ExcelArgument(Name="TTM",Description="The time to maturity associated with the SVI raw parameters. Only ttm, no dates")>] ttm : obj,
                        [<ExcelArgument(Name="Forward",Description="The forward associated with the SVI raw parameters")>] fwd : obj,
                        [<ExcelArgument(Name="Strike",Description="The strike to get the implied vol for")>] strike : obj) =
        validaterawparams a b m r s
            |> Option.bind (fun (a,b,m,r,s) ->  let ttm = ttm |> validateFloat
                                                let fwd = fwd |> validateFloat
                                                let strike = strike |> validateFloat
                                                if ttm.IsNone || fwd.IsNone || strike.IsNone then None
                                                else SVI.impliedvol(a,b,r,m,s) ttm.Value fwd.Value strike.Value)
            |> valueToExcel
 
    [<ExcelFunction(Name="SVI.JWVolatility",Description="Volatility for given strike from SVI Jump-Wing parameters",Category="BJFinanceLib.Volatility",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDF_SVI_JW_Vol ([<ExcelArgument(Name="atmvol",Description="The atm vol parameter in the SVI JW parametrization. Note that this has to be inputted as vol, not variance")>] atm : obj,
                        [<ExcelArgument(Name="atmskew",Description="The atm skew parameter in the SVI JW parametrization.")>] skew : obj,
                        [<ExcelArgument(Name="putskew",Description="The put skew parameter in the SVI JW parametrization.")>] put : obj,
                        [<ExcelArgument(Name="callskew",Description="The call skew parameter in the SVI JW parametrization.")>] call : obj,
                        [<ExcelArgument(Name="minvol",Description="The minimum vol parameter in the SVI JW parametrization. Note that this has to be inputted as vol, not variance")>] minv : obj,
                        [<ExcelArgument(Name="TTM",Description="The time to maturity associated with the SVI raw parameters. Only ttm, no dates")>] ttm : obj,
                        [<ExcelArgument(Name="Forward",Description="The forward associated with the SVI raw parameters")>] fwd : obj,
                        [<ExcelArgument(Name="Strike",Description="The strike to get the implied vol for")>] strike : obj) =
        validatejwparameters atm skew put call minv
            |> Option.bind (fun (atm,skew,put,call,minv) ->  let ttm = ttm |> validateFloat
                                                             let fwd = fwd |> validateFloat
                                                             let strike = strike |> validateFloat
                                                             if ttm.IsNone || fwd.IsNone || strike.IsNone then None
                                                             else
                                                                let rparams = SVI.jumpwingtoraw (atm,skew,put,call,minv) ttm.Value
                                                                if rparams.IsNone then None else SVI.impliedvol rparams.Value ttm.Value fwd.Value strike.Value)
            |> valueToExcel
       
    [<ExcelFunction(Name="SVI.RawToJumpwing",Description="Converts SVI raw parameters into SVI Jump-Wing parameters. The atmf var and min var are returned as volatilities",Category="BJFinanceLib.Volatility",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDF_SVI_RAW_2_J([<ExcelArgument(Name="a",Description="The a parameter in the SVI Raw parametrization of total variance")>] a : obj,
                        [<ExcelArgument(Name="b",Description="The b parameter in the SVI Raw parametrization of total variance")>] b : obj,
                        [<ExcelArgument(Name="r",Description="The rho parameter in the SVI Raw parametrization of total variance")>] r : obj,
                        [<ExcelArgument(Name="m",Description="The mu parameter in the SVI Raw parametrization of total variance")>] m : obj,
                        [<ExcelArgument(Name="s",Description="The sigma parameter in the SVI Raw parametrization of total variance")>] s : obj,
                        [<ExcelArgument(Name="TTM",Description="The time to maturity associated with the SVI raw parameters. Only ttm, no dates")>] ttm : obj) =
        validaterawparams a b m r s
            |> Option.bind (fun (a,b,m,r,s) -> ttm  |> validateFloat
                                                    |> Option.bind (fun t -> if t <= 0. then None else SVI.rawtojumpwing(a,b,r,m,s) t))
            |> Option.bind( fun (vt,phit,pt,ct,vtilde) -> Some(arrayToExcelRow [|sqrt(vt);phit;pt;ct;sqrt(vtilde)|]))
            |> fun x -> match x with
                            | None      -> Array2D.init 1 5 (fun i j -> ExcelError.ExcelErrorNA :> obj)
                            | Some(v)   -> v

    [<ExcelFunction(Name="SVI.JumpwingToRaw",Description="Converts SVI Jump-Wing parameters into SVI Raw parameters. The atmf var and min var numbers have to be inputted as volatilites, not variance",Category="BJFinanceLib.Volatility",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDF_SVI_JW_2_R([<ExcelArgument(Name="atmvol",Description="The atm vol parameter in the SVI JW parametrization. Note that this has to be inputted as vol, not variance")>] atm : obj,
                        [<ExcelArgument(Name="atmskew",Description="The atm skew parameter in the SVI JW parametrization.")>] skew : obj,
                        [<ExcelArgument(Name="putskew",Description="The put skew parameter in the SVI JW parametrization.")>] put : obj,
                        [<ExcelArgument(Name="callskew",Description="The call skew parameter in the SVI JW parametrization.")>] call : obj,
                        [<ExcelArgument(Name="minvol",Description="The minimum vol parameter in the SVI JW parametrization. Note that this has to be inputted as vol, not variance")>] minv : obj,
                        [<ExcelArgument(Name="TTM",Description="The time to maturity associated with the JW parameters. Only ttm, no dates")>] ttm : obj) =
        validatejwparameters atm skew put call minv 
            |> Option.bind(fun (atm,skew,put,call,minv) ->  ttm |> validateFloat
                                                                |> Option.bind (fun t -> if t <= 0. then None else SVI.jumpwingtoraw(atm,skew,put,call,minv) t))
            |> Option.bind( fun (a,b,r,m,s) -> Some(arrayToExcelRow [|a;b;r;m;s|]))
            |> fun x -> match x with
                            | None      -> Array2D.init 1 5 (fun i j -> ExcelError.ExcelErrorNA :> obj)
                            | Some(v)   -> v

    [<ExcelFunction(Name="SVI.Calibrate",Description="Calibrates raw SVI parameters on given input strieks and vols",Category="BJFinanceLib.Volatility",IsThreadSafe=true)>]
    let UDF_SVI_CALIB([<ExcelArgument(Name="Strikes",Description="Input strikes to calibrate on")>] strikes : obj [],
                      [<ExcelArgument(Name="Vols",Description="Input vols to calibration on")>] vols : obj [],
                      [<ExcelArgument(Name="TTM",Description="Time to maturity")>] ttm : obj,
                      [<ExcelArgument(Name="Forward",Description="Forward used in calibration. Optional, default = 1. (in which case your strikes will have to be percentage strikes)")>] fwd : obj,
                      [<ExcelArgument(Name="VegaWeighted",Description="Boolean that determines whether calibration occurs on a vega-weighted basis (if true), or equally weighted basis (if false). Default=False")>] vegaWeighted : obj) =
        let strikes = array1DAsArray validateFloat strikes
        let vols = array1DAsArray validateFloat vols
        let ttm = validateFloat ttm |> FSharpx.Option.getOrElse -1. 
        let fwd = validateFloat fwd |> FSharpx.Option.getOrElse 1.
        let vegaWeighted = validateBool vegaWeighted |> FSharpx.Option.getOrElse false
        if Array.length strikes = Array.length vols && ttm > 0. && fwd > 0. then
            let weightfunc = 
                if vegaWeighted then
                    let maxvega = BJExcelLib.Finance.Blackscholes.vega false (Call) fwd fwd ttm 0. 0. 0.30 |> FSharpx.Option.getOrElse 1. // Not exactly equal to maxvega, but close enough for scaling
                    fun (strike,vol) -> BJExcelLib.Finance.Blackscholes.vega false (Call) fwd strike ttm 0. 0. vol |> FSharpx.Option.getOrElse 0. |> fun v -> v/maxvega
                else fun (strike,vol) -> 1.
            Array.zip strikes vols  |> Array.filter (fun (s,v) -> s.IsSome && v.IsSome)
                                    |> Array.map (fun (s,v) -> s.Value, v.Value)
                                    |> Array.map (fun (s,v) -> weightfunc (s,v),s,v)
                                    |> fun points -> SVI.calibrate (SVI_CALIB_CHECK_FOR_ARB,SVI_CALIB_MIN_SIGMA) ttm fwd points
                                    |> fst
                                    |> fun (a,b,r,m,s) -> [|a;b;r;m;s|]
        else [| 0.;0.;0.;0.;0.|]
        |> arrayToExcelRow