namespace BJExcelLib.UDF

/// Contains UDFs for american option approximations functions
module public American=
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


    [<ExcelFunction(Name="JZ.Value",Description="Value of an american option in the Black-Scholes model with Ju-Zhong approximation",Category="BJExcelLib.OptionPricing",IsExceptionSafe=false,IsThreadSafe=true)>]
    let UDFJZValue ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> JuZhong.value marg flag S K T r b v)
            |> valueToExcel

(*TODO: these need to be checked first
    [<ExcelFunction(Name="JZ.Delta",Description="Delta of an american option in the Black-Scholes model with Ju-Zhong approximation",Category="BJExcelLib.OptionPricing",IsExceptionSafe=false,IsThreadSafe=true)>]
    let UDFJZDelta ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> JuZhong.delta marg flag S K T r b v)
            |> valueToExcel

    [<ExcelFunction(Name="JZ.Gamma",Description="Delta of an american option in the Black-Scholes model with Ju-Zhong approximation",Category="BJExcelLib.OptionPricing",IsExceptionSafe=false,IsThreadSafe=true)>]
    let UDFJZGamma ([<ExcelArgument(Name="Spot",Description = "Spot of the underlying")>] S : obj,
                    [<ExcelArgument(Name="Strike",Description = "Option strike")>] K : obj,
                    [<ExcelArgument(Name="vol",Description = "Volatility as a percentage")>] v : obj,
                    [<ExcelArgument(Name="TTM",Description = "Time to maturity. No dates")>] T : obj,
                    [<ExcelArgument(Name="rate",Description = "Interest rate as a percentage. Optional, default = 0%")>] r : obj,
                    [<ExcelArgument(Name="costcarry",Description = "Cost of carry as a percentage. Note div yield = rate-costofcarry Optional, default = r")>] b : obj,
                    [<ExcelArgument(Name="Flag",Description = "C for call, P for Put. Optional, default inferred from spot and strike")>] flag : obj,
                    [<ExcelArgument(Name="margined",Description="Indicates whether the option premium gets margined or not. Optional,default= false")>] marg : obj) =
        validateBS S K v r b T flag marg
            |> Option.bind (fun (S,K,v,r,b,T,flag,marg) -> JuZhong.gamma marg flag S K T r b v)
            |> valueToExcel
            *)
