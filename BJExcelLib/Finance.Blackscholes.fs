namespace BJExcelLib.Finance

type OptionType =
    | Call
    | Put

/// Can't have a finance library without blackscholes formulae
module public Blackscholes =

    open BJExcelLib.Math.NormalDistribution

    // Parameters that determine how searching for implied is done.
    [<Literal>]
    let private VOL_ITER = 100;         // Max iterations to search for volatility
    [<Literal>]
    let private LOW_VOL = 0.001         // Don't search for vols below this
    [<Literal>]
    let private HIGH_VOL = 5.           // Don't search for vols above this
    [<Literal>]
    let private VOL_PRECISION = 0.0001  // Search until convergence in option price down to 1 basis point of spot.
                                        // Basically this will fail to converge well for very low premium options (think deep otm with low vol)
                                        // but there you can argue that uncertainty about vega will be large anyway (and is unlikely as well).

    /// Root finding algorithm used to find volatilities. Switch here and its switched everywhere in this module
    let private rootfinder spot f = MathNet.Numerics.FindRoots.bisection VOL_ITER (VOL_PRECISION*spot) LOW_VOL HIGH_VOL f

    /// returns (flag,df,d1,d2)
    let private getstuff margined flag S K t r b v =
        let inputOk = S > 0. && K >=0. && v > 0. && t > 0.
        if not inputOk then None else
            let flag = if flag = Call then 1. else -1.
            let df = if margined then 1. else exp(-r*t)
            let vsqrt = v * sqrt t
            let d1 = (log(S/K)+b*t)/vsqrt + 0.5*vsqrt
            Some(flag,df,d1,d1-vsqrt)

    // Forward of an underlying. Not really specific to BS world, but got to put it somewhere
    let public forward  S t b =
        if S <= 0. || t < 0. then None
        else Some(S*exp(b*t))

    /// Value of a vanilla european option using generalized BS
    /// Apart from the usual parameters, there is a parameter
    /// margined here which specifies whether the option value gets
    /// margined (this happens for example with Brent options on the ICE)
    let public value margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,d1,d2)   ->  df*flag*( S*exp(b*t)*cnd(flag*d1)-K*cnd(flag*d2) ) |>Some
        

    /// Value, delta and vega of a vanialla european option using generalized BS
    /// Apart from the usual parameters, there is a parameter
    /// margined here which specifies whether the option value gets
    /// margined (this happens for example with Brent options on the ICE)
    let public valuedeltavega margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,d1,d2)   ->  let expr1 = exp(b*t)
                                        let expr2 = cnd(flag*d1)
                                        let value = df*flag*( S*expr1*expr2 - K*cnd(flag*d2) )
                                        let delta = df*flag*expr1*expr2 
                                        let vega = df*K*nd(d2)*sqrt(t)*0.01
                                        Some(value,delta,vega)
 
    /// Delta of a vanilla european option using generalized BS
    let public delta margined flag S K t r b v  =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,d1,d2)   ->  Some(flag*df*exp(b*t)*cnd(flag*d1))
    
    /// Vega of a vanilla european option using generalized BS    
    let public vega margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(_,df,_,d2)       ->  Some(df*K*nd(d2)*sqrt(t)*0.01)

    /// Daily theta of a vanilla european option using generalized BS
    let public theta margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,d1,d2)   ->  let fwd = S*exp(b*t)
                                        Some(df*(-fwd*(nd(d1)*v/(2.*sqrt(t))+flag*(r-b)*cnd(flag*d1))-flag*r*K*cnd(flag*d2))/365.25)

    /// Gamma of a vnailla european option using generalized BS
    let public gamma margined flag S K t r b v = 
        match getstuff margined flag S K t r b v with
            | None              ->  None
            | Some(_,df,d1,_)   ->  Some(df*exp(b*t)*nd(d1)/(S*v*sqrt(t)))

    /// Rho per basis point of a vanilla european option using generalized BS
    let public rho margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,_,d2)    ->  Some(flag*K*t*df*cnd(flag*d2)*0.0001)

    /// Dual delta (sensitivity to strike) of a vanilla european option using generalized BS
    let public dualdelta margined flag S K t r b v  =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,_,d2)    ->  Some(-flag*df*cnd(flag*d2))
    
    /// dVegadSpot = dDeltadVol = Vanna of a vanilla european option in a generalied BS setting
    let public dvegadspot  margined flag S K t r b v  =
        match getstuff margined flag S K t r b v with
            | None              ->  None
            | Some(_,df,d1,d2)  ->  Some(-exp(b*t)*df*nd(d1)*d2/v)

    /// dVegadVol = Volga of a vanilla european option in a generalized BS world
    let public dvegadvol margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None              ->  None
            | Some(_,df,d1,d2)  ->  Some(df*K*nd(d2)*sqrt(t)*0.01*d1*d2/v)

    /// ddeltadtim = Charm = Delta bleed of a vanilla european option in a generalized BS world
    let public ddeltadtime margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                  ->  None
            | Some(flag,df,d1,d2)   ->  let bt = b*t
                                        let vsqrt = v*sqrt(t)
                                        Some(df*exp(bt)*((r-b)*flag*cnd(flag*d1)-nd(d1)*(2.*bt-d2*vsqrt)/(2.*t*vsqrt)))

    /// Implied volatility for an european option with a given premium in a black-scholes world
    let public impliedvol margined flag S K t r b premium =
        let price vol = (value margined flag S K t r b vol) |> Option.bind (fun x -> Some(x-premium))
        match price LOW_VOL, price HIGH_VOL with
            | None, _ | _, None -> None
            | Some(x), Some(y)  ->  if x > 0. || y < 0. then None
                                    else let price = price >> Option.get
                                         rootfinder S price
                                         // todo : improve by using stuff from P. Jaeckel's by implication or from the things Axel Vogt has written on numerics.
    

    /// Value of a european straddle using generalized BS
    let public straddlevalue margined S K t r b v =
        match getstuff margined Call S K t r b v with
            | None              ->  None
            | Some(_,df,d1,d2)  ->  df*( S*exp(b*t)*(cnd(d1)-cnd(-d1))-K*(cnd(d2)-cnd(-d2)) ) |>Some

    /// Delta of a european straddle using generalized BS
    let public straddledelta margined S K t r b v  =
        match getstuff margined Call S K t r b v with
            | None              ->  None
            | Some(_,df,d1,d2)  ->  df*exp(b*t)*(cnd(d1)-cnd(-d1)) |> Some

    /// Vega of a european straddle using generalized BS    
    let public straddlevega margined S K t r b v =
        match getstuff margined Call S K t r b v with
            | None              ->  None
            | Some(_,df,_,d2)   ->  Some(2.*df*K*nd(d2)*sqrt(t)*0.01)


    /// Implied volatility for an european option with a given premium in a black-scholes world
    let public straddleimpliedvol margined S K t r b premium =
        let price vol = (straddlevalue margined S K t r b vol) |> Option.bind (fun x -> Some(x-premium))
        match price LOW_VOL, price HIGH_VOL with
            | None, _ | _, None -> None
            | Some(x), Some(y)  ->  if x > 0. || y < 0. then None
                                    else let price = price >> Option.get
                                         rootfinder S price