namespace BJExcelLibTests

open System
open Microsoft.VisualStudio.TestTools.UnitTesting
open BJExcelLib.Finance


[<TestClass>]
type BlackScholesTest() =
 
    let TOLERANCE_PUTCALLPARITY = 0.0001    // Tolerance for put-call parity errors (as percentage of spot)
    let TOLERANCE_BSVALUES = 0.0001         // Tolerance for option values
    let TOLERANCE_BSGREEKS = 0.001          // Tolerance on numerical greek values
    let TOLERANCE_VOL = 0.0005;             // Tolerance on search for vol


    /// All blackscholes functions that have the same input types. Basically all of them except for the forward and inverse search functions.
    let mostBSfuncs = [ Blackscholes.value; Blackscholes.delta; Blackscholes.vega; Blackscholes.gamma; Blackscholes.rho; Blackscholes.theta; Blackscholes.dvegadspot; Blackscholes.dvegadspot; Blackscholes.ddeltadtime]
    let straddlefuncs = [Blackscholes.straddlevalue; Blackscholes.straddledelta; Blackscholes.straddlevega]
    // Checks whether all blackscholes functions return none if time to maturity is negative. Splitting out per function would be more atomic, but too much work for the gain.
    [<TestMethod>]
    member x.``All Black-scholes functions return None if time to maturity is negative`` () = 
        let marg, flag,S,K,t,r,b,v = false, Call,100.,100.,-1.,0.05,0.03,0.20
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f -> acc && Option.isNone (f marg flag S K t r b v)) true
        let straddleresult = straddlefuncs |> List.fold (fun acc f -> acc && Option.isNone (f marg S K t r b v)) true
        let forwardResult = Option.isNone (Blackscholes.forward S t b)                          
        Assert.AreEqual(true, mostfuncresult && forwardResult && straddleresult)

    // Checks whether all blackscholes functions return none if the spot is negative. Splitting out per function would be more atomic, but too much work for the gain.
    [<TestMethod>]
    member x.``All Black-scholes functions return None if the spot is negative`` () = 
        let marg, flag,S,K,t,r,b,v = false, Call,-100.,100.,1.,0.05,0.03,0.20
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f -> acc && Option.isNone (f marg flag S K t r b v)) true
        let straddleresult = straddlefuncs |> List.fold (fun acc f -> acc && Option.isNone (f marg S K t r b v)) true
        let forwardResult = Option.isNone (Blackscholes.forward S t b)                          
        Assert.AreEqual(true, mostfuncresult && forwardResult && straddleresult)

    // Checks whether all relevant blackscholes functions return none if the strike. Splitting out per function would be more atomic, but too much work for the gain.
    [<TestMethod>]
    member x.``All Black-scholes functions return None if the strike is negative`` () = 
        let marg, flag,S,K,t,r,b,v = false, Call,100.,-100.,1.,0.05,0.03,0.20
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f -> acc && Option.isNone (f marg flag S K t r b v)) true
        let straddleresult = straddlefuncs |> List.fold (fun acc f -> acc && Option.isNone (f marg S K t r b v)) true                     
        Assert.AreEqual(true, mostfuncresult && straddleresult)

    // Checks whether all blackscholes functions return none if the spot is zero. Splitting out per function would be more atomic, but too much work for the gain.
    [<TestMethod>]
    member x.``All Black-scholes functions return None if the spot is zero`` () = 
        let marg, flag,S,K,t,r,b,v = false, Call,0.,100.,1.,0.05,0.03,0.20
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f -> acc && Option.isNone (f marg flag S K t r b v)) true
        let straddleresult = straddlefuncs |> List.fold (fun acc f -> acc && Option.isNone (f marg S K t r b v)) true
        let forwardResult = Option.isNone (Blackscholes.forward S t b)                          
        Assert.AreEqual(true, mostfuncresult && forwardResult && straddleresult)

    // Checks whether all blackscholes functions return none if volatility is negative. Splitting out per function would be more atomic, but too much work for the gain.
    [<TestMethod>]
    member x.``All relevant Black-scholes functions return None if the volatility is negative`` () = 
        let marg, flag,S,K,t,r,b,v = false, Call,100.,100.,1.,0.05,0.03,-0.20
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f -> acc && Option.isNone (f marg flag S K t r b v)) true 
        let straddleresult = straddlefuncs |> List.fold (fun acc f -> acc && Option.isNone (f marg S K t r b v)) true                      
        Assert.AreEqual(true, mostfuncresult && straddleresult)

    // Checks whether the implementation of option valuation is consistent with the forward function up to specified tolerance.
    [<TestMethod>]
    member x.``Blackscholes option values preserve put call parity up to required tolerance`` () =
        let result = [(100., 100., 1.,  0.05,  0.03, 0.20);
                      (100., 80.,  2.,  0.05,  0.03, 0.80);
                      (100., 200., 1.2,-0.02,  0.01, 0.30);
                      (100., 110., 0.3, 0.02, -0.01, 0.35)]
                        |> List.fold (fun acc  (spot,strike,ttm,rate,cc,vol) ->
                                        let call = Blackscholes.value false Call spot strike ttm rate cc vol
                                        let put = Blackscholes.value false Put spot strike ttm rate cc vol
                                        let forward = Blackscholes.forward spot ttm cc
                                        if call.IsNone || put.IsNone || forward.IsNone then false
                                        else ((call.Value-put.Value)*exp(rate*ttm)+strike-forward.Value)/spot |> abs |> fun x -> x < TOLERANCE_PUTCALLPARITY) true
        Assert.AreEqual(true,result)

    [<TestMethod>]
    member x.``Blackscholes option values and greeks when the premium are margined differ from non-margined options by a multiplicative factor equal to the discountfactor``() =
        let flag,S,K,t,r,b,v = Call,100.,100.,0.3,0.05,0.03,0.20
        let df = exp(-r*t)
        let mostfuncresult = mostBSfuncs |> List.fold(fun acc f ->  let unmargined  = f false flag S K t r b v
                                                                    let margined    = f true flag S K t r b v
                                                                    if unmargined.IsNone || margined.IsNone then false
                                                                    else acc && ((unmargined.Value/(margined.Value*df)) - 1. |> abs |> fun x -> x < 0.000000001 )) true
        let straddleresult = straddlefuncs |> List.fold(fun acc f ->let unmargined  = f false S K t r b v
                                                                    let margined    = f true S K t r b v
                                                                    if unmargined.IsNone || margined.IsNone then false
                                                                    else acc && ((unmargined.Value/(margined.Value*df)) - 1. |> abs |> fun x -> x < 0.000000001)) true        

        Assert.AreEqual(true,mostfuncresult && straddleresult)

    [<TestMethod>]
    member x.``Blackscholes value is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.value marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp)/S < TOLERANCE_BSVALUES)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes delta is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.delta marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp) < TOLERANCE_BSGREEKS)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes vega is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.vega marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp) < TOLERANCE_BSGREEKS)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes gamma is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.gamma marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp) < TOLERANCE_BSGREEKS)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes theta is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.theta marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp) < TOLERANCE_BSGREEKS)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes rho is correct``() =
        let test = [((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : get correct values for expected results and increase number of tests
                    ((false,Call,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,flag,S,K,t,r,b,v),exp) -> Blackscholes.rho marg flag S K t r b v
                                                                                |> Option.map (fun res -> abs(res-exp) < TOLERANCE_BSGREEKS)
                                                                                |> fun x -> if x.IsNone then false else x.Value && acc) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes straddle price is equal to sum of put and call prices``() =
        let test = [((false,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : increase number of tests
                    ((false,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,S,K,t,r,b,v),exp) ->   let call = Blackscholes.value marg Call S K t r b v
                                                                            let put = Blackscholes.value marg Put S K t r b v
                                                                            let straddle = Blackscholes.straddlevalue marg S K t r b v
                                                                            if call.IsNone || put.IsNone || straddle.IsNone then false
                                                                            else (call.Value+put.Value-straddle.Value)/S < TOLERANCE_BSVALUES) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes straddle vega is equal to sum of put and call prices``() =
        let test = [((false,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : increase number of tests
                    ((false,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,S,K,t,r,b,v),exp) ->   let call = Blackscholes.vega marg Call S K t r b v
                                                                            let put = Blackscholes.vega marg Put S K t r b v
                                                                            let straddle = Blackscholes.straddlevega marg S K t r b v
                                                                            if call.IsNone || put.IsNone || straddle.IsNone then false
                                                                            else (call.Value+put.Value-straddle.Value)/S < TOLERANCE_BSVALUES) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes straddle delta is equal to sum of put and call prices``() =
        let test = [((false,100.,100.,1., 0.05, 0.02, 0.30),1.); //todo : increase number of tests
                    ((false,100.,100.,1., 0.05, 0.02, 0.30),1.)]
                        |> List.fold (fun acc ((marg,S,K,t,r,b,v),exp) ->   let call = Blackscholes.delta marg Call S K t r b v
                                                                            let put = Blackscholes.delta marg Put S K t r b v
                                                                            let straddle = Blackscholes.straddledelta marg S K t r b v
                                                                            if call.IsNone || put.IsNone || straddle.IsNone then false
                                                                            else (call.Value+put.Value-straddle.Value)/S < TOLERANCE_BSVALUES) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes implied vol search returns expected vol up to required tolerance``() =
        let test = [(false,Call,100.,100.,  1.,   0.10, 0.00,  0.30);
                    (false,Put, 25.,30.,    2.,   0.05, 0.02,  0.25);
                    (true,Call, 10.,5.,     4.,   0.01, 0.03,  0.10);
                    (false,Put, 1.,0.7,     0.5,  0.02, 0.02,  0.15);
                    (true,Call, 0.1,0.12,  2.,   0.00, 0.00,   1.33);
                    (false,Put, 100.,100.,  10.,  0.00, -0.05, 0.10);
                    (true,Call, 100.,100.,  0.05, 0.03, 0.04,  0.55);
                    (false,Put, 100.,100.,  0.20, 0.02, 0.,    0.08);
                    (false,Call, 100.,100., 7.,    0.07, 0.01,  0.36);]
                        |> List.fold (fun acc (marg,flag,S,K,t,r,b,v) ->    let vega = Blackscholes.vega marg flag S K t r b v
                                                                            Blackscholes.value marg flag S K t r b v
                                                                                |> Option.bind (fun premium ->  Blackscholes.impliedvol marg flag S K t r b premium
                                                                                                                    |> Option.map (fun r -> abs(r-v) < TOLERANCE_VOL || vega.Value / premium < 0.00001 || premium / S < 0.00001 )) // Vega converged or worthless option and/or worthless time value in option
                                                                                |> fun x -> if x.IsNone then false else acc && x.Value) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Blackscholes straddle implied vol search returns expected vol up to required tolerance``() =
        let test = [(false,100.,100.,1., 0.05, 0.02, 0.30);
                    (false,100., 40.,1., 0.05, 0.02, 0.35);
                    (false,10.,30.,  1., 0.05, 0.02, 0.80);
                    (false,100.,150.,2.5, 0.05, 0.02, 1.30);
                    (false,1., 1.1,0.2, 0.05, 0.02, 0.26);
                    (true,100.,95.,0.06, 0.05, 0.02, 0.28);
                    (false,100.,101.,8., 0.05, 0.02, 0.15);
                    (true,100.,99.,0.5, 0.05, 0.02, 0.05);
                    (false,100.,103.,3., 0.05, 0.02, 0.25)
                    ]
                        |> List.fold (fun acc (marg,S,K,t,r,b,v) ->  Blackscholes.straddlevalue marg S K t r b v
                                                                        |> Option.bind (fun premium ->  let vega = Blackscholes.straddlevega marg S K t r b v
                                                                                                        Blackscholes.straddleimpliedvol marg S K t r b premium
                                                                                                            |> Option.map (fun r -> abs(r-v) < TOLERANCE_VOL || vega.Value / premium < 0.00001 || premium / S < 0.00001 )) // Vega converged or worthless option and/or worthless time value in option
                                                                        |> fun x -> if x.IsNone then false else acc && x.Value) true
        Assert.AreEqual(true,test)

