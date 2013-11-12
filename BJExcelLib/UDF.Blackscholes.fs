namespace BJExcelLib.UDF

/// Contains UDFs for black-scholes functions
module public Blackscholes=
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.Finance

    /// Validates black-scholes input parameters
    let private validateBS S K v r b T flag marg =
        let S = S |> validateFloat |> Option.bind (fun x -> if x < 0. then None else Some(x))
        let K = K |> validateFloat |> Option.bind (fun x -> if x < 0. then None else Some(x))
        let v = v |> validateFloat |> Option.bind (fun x -> if x <= 0. then None else Some(x))
        let r = r |> validateFloat |> FSharpx.Option.getOrElse 0.
        let b = b |> validateFloat |> FSharpx.Option.getOrElse r
        let T = T |> validateFloat |> Option.bind (fun x -> if x <= 0. then None else Some(x))
        let flag = flag |> validateString
                        |> Option.bind (fun (x : string)-> let z = (x.ToUpper())
                                                           if z.Length = 0 then None
                                                           else
                                                                let z = z.[0]
                                                                if z = 'C' then Some(Call) elif z = 'P' then Some(Put) else None )
                        |> FSharpx.Option.getOrElse (if S <= K then Call else Put)
        let marg = marg |> validateBool |> FSharpx.Option.getOrElse false
        if S.IsSome && K.IsSome && T.IsSome && v.IsSome then Some(S.Value,K.Value,v.Value,r,b,T.Value,flag,marg) else None

    //  Generalized BS UDFs
    [<ExcelFunction(Name="Forward",Description="Forward of an underlying from generalized Black-Scholes input parameters",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFForward ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = 0")>] b : obj) =
        validateBS S S 0.1 0. b T Call false
            |> Option.bind (fun (S,_,_,_,b,T,_,marg) -> Blackscholes.forward S T b)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Value",Description="Value of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSValue ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.value marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Delta",Description="Delta of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSDelta ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.delta marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Vega",Description="vega of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSvega  ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.vega marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Gamma",Description="gamma of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSgamma ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.gamma marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.DualDelta",Description="dualdelta of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSdualdelta ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                        [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                        [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                        [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                        [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                        [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                        [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                        [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.dualdelta marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Rho",Description="rho of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSrho   ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.rho marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.Theta",Description="theta of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBStheta ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.theta marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.dVegadSpot",Description="dvegadspot of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSdvegadspot([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                        [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                        [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                        [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                        [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                        [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                        [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                        [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.dvegadspot marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.dVegadVol",Description="dvegadvol of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSdvegadvol ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                        [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                        [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                        [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                        [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                        [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                        [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                        [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.dvegadvol marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.dDeltadTime",Description="ddeltadtime of an european option in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSddeltadtime   ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                            [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.ddeltadtime marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BS.ValueDeltaVega",Description="Value, delta and vega of an european option in the Black-Scholes model as an array",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSvaldeltavega  ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                            [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.valuedeltavega marg flag S K T r b v)
            |> Option.map (fun (v,d,vega)-> [|v;d;vega|])
            |> FSharpx.Option.getOrElse Array.empty
            |> arrayToExcelRow

    [<ExcelFunction(Name="BS.ImpliedVol",Description="Implied volatility of an european option in the Black-Scholes model for a given premium. Works ok except for deep OTM options with low vol",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSImplied   ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                        [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                        [<ExcelArgument(Name="Premium",Description = "Option premium")>] prem : obj,
                        [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                        [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                        [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                        [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                        [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K 0.10 r b T flag marg // Just use some dummy value for v
            |> Option.bind (fun (S,K,_,r,b,T,flag,marg) ->  let prem = prem |> validateFloat |> Option.bind (fun x -> if x <= 0. then None else Some(x))
                                                            if prem.IsNone then None else Blackscholes.impliedvol marg flag S K T r b prem.Value)
            |> valueToExcel


    // UDF for Black-Scholes on futures
    [<ExcelFunction(Name="BSFuture.Value",Description="Value of an european option on a future in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let USDBSFutureValue   ([<ExcelArgument(Name="Spot",Description = "Price of the future")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r 0. T flag marg
            |> Option.bind (fun (S,K,v,r,_,T,flag,marg) -> Blackscholes.value marg flag S K T r 0. v)
            |> valueToExcel

    [<ExcelFunction(Name="BSFuture.Delta",Description="Delta of an european option on a future in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let USDBSFutureDelta   ([<ExcelArgument(Name="Spot",Description = "Price of the future")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r 0. T flag marg
            |> Option.bind (fun (S,K,v,r,_,T,flag,marg) -> Blackscholes.delta marg flag S K T r 0. v)
            |> valueToExcel

    [<ExcelFunction(Name="BSFuture.Vega",Description="vega of an european option on a future in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let USDBSFuturevega([<ExcelArgument(Name="Spot",Description = "Price of the underlying")>] S : obj,
                        [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                        [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                        [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                        [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                        [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                        [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r 0. T flag marg
            |> Option.bind (fun (S,K,v,r,_,T,flag,marg) -> Blackscholes.vega marg flag S K T r 0. v)
            |> valueToExcel

    [<ExcelFunction(Name="BSFuture.ImpliedVol",Description="Implied volatility of an option on a future in the Black-Scholes model for a given premium. Works ok except for deep OTM options with low vol",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSFutureImplied ([<ExcelArgument(Name="Spot",Description = "Price of the future")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="Premium",Description = "Option premium")>] prem : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K 0.10 r 0. T flag marg // Just use some dummy value for v
            |> Option.bind (fun (S,K,_,r,_,T,flag,marg) ->  let prem = prem |> validateFloat |> Option.bind (fun x -> if x <= 0. then None else Some(x))
                                                            if prem.IsNone then None else Blackscholes.impliedvol marg flag S K T r 0. prem.Value)
            |> valueToExcel

    // UDFs for Black-scholes on straddles
    [<ExcelFunction(Name="BSStraddle.Value",Description="Value of an european straddle in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSStraddleValue ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T Call marg // Just use some dummy value for flag
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.straddlevalue marg S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BSStraddle.Delta",Description="Delta of an european straddle in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSStraddleDelta ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T Call marg // Just use some dummy value for flag
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.straddledelta marg S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BSStraddle.Vega",Description="vega of an european straddle in the Black-Scholes model",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSStraddlevega  ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                            [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                            [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                            [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                            [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                            [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                            [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T Call marg // Just use some dummy value for flag
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> Blackscholes.straddlevega marg S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="BSStraddle.ImpliedVol",Description="Implied volatility of an european straddle in the Black-Scholes model for a given premium. Works ok except for deep OTM options with low vol",Category="BJExcelLib.OptionPricing",IsExceptionSafe=true,IsThreadSafe=true)>]
    let UDFBSStraddleImplied   ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                                [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                                [<ExcelArgument(Name="Premium",Description = "Option premium")>] prem : obj,
                                [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                                [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                                [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                                [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K 0.10 r b T Call marg // Just use some dummy value for v & flag
            |> Option.bind (fun (S,K,_,r,b,T,flag,marg) ->  let prem = prem |> validateFloat |> Option.bind (fun x -> if x <= 0. then None else Some(x))
                                                            if prem.IsNone then None else Blackscholes.straddleimpliedvol marg S K T r b prem.Value)
            |> valueToExcel