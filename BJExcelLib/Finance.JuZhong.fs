namespace BJExcelLib.Finance

/// American option approximation according to Ju & Zhong
module JuZhong =
    open Blackscholes
    open BJExcelLib.Math.NormalDistribution

    /// returns (flag,hAh,S*,lambda(h),X,X',X'',d1(S*)) as an option type
    let private getstuff margined flag S K t r cc v =
        let inputOk = S > 0. && K >=0. && v > 0. && t > 0.
        if not inputOk then None else
        let flag' = if flag = Call then 1. else -1.
        if margined then (flag',0.,1.,0.,0.,0.,0.,0.) |> Some // if margined, return paramters that will default to europena option value
        else
            let alpha = 2.*r/(v*v)
            let beta = 2.*cc/(v*v) - 1. // compared to JZ paper, this is beta-1
            let h = 1.- exp(-r*t)
            let mdivy = cc-r //minus one times the div yield
            let commonExp = exp(mdivy*t)
            let subexpr0 = v*v*t
            let subexpr1 = beta*beta + 8./subexpr0
            let subexpr2 = sqrt(subexpr1)
            let subexpr3 = sqrt(beta*beta + 4.*alpha/h)
            let lambda, lambda', b =
                if r = 0. && flag' = 1. then
                    0.5*(flag'*subexpr2-beta), 0. (*not used*), -2./(subexpr0*subexpr0*subexpr1)
                else
                    let lambda' = -flag'*alpha/(h*h*subexpr3)
                    let lambda = 0.5*(flag'*subexpr3-beta)
                    lambda, lambda', (1.-h)*alpha*lambda'/(2.*(2.* lambda + beta))

            let d1Func sStar = (log(sStar/K) + (cc + 0.5*v*v)*t)/(v*sqrt(t))
            let sStarFunc S' = -flag' + flag'*commonExp*cnd(flag'*(d1Func S')) + lambda*(flag'*(S'-K) - (Blackscholes.value margined flag S' K t r cc v).Value)/S' 
            let d f x =
                let step = sqrt BJExcelLib.Math.Constants.machine_epsilon
                (f (x + step) - f(x - step))/2./step
            
            let iter = 300
            let acc = (0.0001 * (min S K))
            let lb = (0.5 * (min S K))
            let ub = (5. * (max S K))
            let sStar = MathNet.Numerics.FindRoots.newtonRaphsonGuess iter acc S sStarFunc (d sStarFunc)
            if sStar.IsNone then None
            else
                let sStar = sStar.Value
                let hAh = Blackscholes.value margined flag sStar K t r cc v
                        |> Option.map (fun ve -> flag'*(sStar-K) - ve)
                if hAh.IsNone then None
                else
                    let hAh = hAh.Value
                    let d1S' = d1Func sStar
                    let d2S' = d1S' - v*sqrt(t)
                    let c =
                        if r = 0. && flag' = 1. then
                            -flag'/subexpr2*(2./subexpr0 + 2.*b + (sStar*commonExp/(hAh*v))*(nd(d1S')/sqrt(t) + 2.*flag'*mdivy/v))
                        else
                            let subexpr4 = 2.*lambda+beta
                            let sens = (sStar*exp(cc*t)/r)*(nd(d1S')*v/(2.*sqrt(t))+flag'*mdivy*cnd(flag'*d1S'))+flag'*K*cnd(flag'*d2S')
                            (h-1.)*alpha/subexpr4*(lambda'/subexpr4 + 1./h + sens/hAh)

                    let logSS' = log(S/sStar)
                    let X, X', X'' = b*logSS'*logSS' + c * logSS',
                                     (2.*b*logSS' + c)/S,
                                     (2.*b*(1.-logSS')-c)/(S*S)
            
                    (flag',hAh,sStar,lambda,X,X',X'',d1S') |> Some

    /// Value of an american option in the Ju & Zhong approximation.
    let public value  margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                              ->  None
            | Some(flag',hAh,sStar,lh,X,_,_,_)  ->  if flag'*(sStar-S) <= 0. then flag'*(S-K) |> Some
                                                    else Blackscholes.value margined flag S K t r b v
                                                            |> Option.map (fun ve -> ve + hAh*((S/sStar)**lh)/(1.-X))
    
    // TODO: these need to be checked
    /// Delta of an american option in the JuZhong approximation
    let public delta margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                              ->  None
            | Some(flag',hAh,sStar,lh,X,X',_,d1)->  Blackscholes.value margined flag sStar K t r b v
                                                        |> Option.map (fun veS' ->  flag'*exp((b-r)*t)*cnd(flag'*d1) + (lh/(S*(1.-X))+X'/((1.-X)*(1.-X)))*(flag'*(sStar-K)-veS')*((S/sStar)**lh))
    
    /// Gamma of an american option in the Ju & Zhong approximation
    let public gamma margined flag S K t r b v =
        match getstuff margined flag S K t r b v with
            | None                              ->  None
            | Some(flg',hAh,sStr,lh,X,X',X'',d1)->  Blackscholes.value margined flag sStr K t r b v
                                                        |> Option.map (fun ve -> flg'*exp((b-r)*t)*nd(flg'*d1) +
                                                                                 (flg'*(sStr-K) - ve)*((S/sStr)**lh) *
                                                                                 ((X'' + (2.*X'*(lh/S+X'/(1.-X))))/((1.-X)*(1.-X)) + (lh*lh-lh)/(S*S*(1.-X))))
                                                                                                                                                    
