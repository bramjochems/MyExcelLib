namespace BJExcelLib.Finance

/// Contains functions relating to Gatheral's SVI function
module SVI =
    open BJExcelLib.Math.Optimization
    open MathNet.Numerics.LinearAlgebra.Double

    /// Checks some bounds on SVI raw parameters at for which there surely is arbitrage. Note that
    /// if this function indicates that bounds aren't exceeded, there still can be arbitrage. However,
    /// it still is worthwhile checking for these. 
    let public rawparameterssane (a,b,rho,m,sigma) ttm =
        ttm > 0. && sigma > 0. && abs(rho) <= 1. && (a+b*sigma*sqrt(1.-rho*rho)) >= 0. && (b*(1.+abs(rho)) < 4./ttm)

    /// Computes total variance using the SVI function given SVI Raw parameters a time to maturity, a forward and a strike as input
    let public totalvariance (a,b,rho,m,sigma) fwd (strike : float) =
        if strike <= 0. || fwd <= 0. then None
        else
            let rho = max -1. (min 1. rho)
            let k = log(strike/fwd)-m
            let res = a + b*(rho*k+sqrt(k*k+sigma*sigma))
            if res >= 0. then Some(res) else None

    /// Computes implied vol using the SVI function given SVI Raw parameters a time to maturity, a forward and a strike as input
    /// Notice that the parameters are used to calculate total variance which then gets translated into implied volatility. This
    /// is different from some implementations where the parameters model implied variance instead of implied total variance. The
    /// difference on the parameters is just a factor T
    let public impliedvol (a,b,rho,m,sigma) ttm fwd strike =
        totalvariance (a,b,rho,m,sigma) fwd strike |> Option.map(fun x -> sqrt(x/ttm))

    /// Gives the gradient of implied vol w.r.t a,b,rho, m and sigma. Usefol for some numerical
    /// optimization schemes
    /// This funciton is easily obtained by seeing impliedvol as sqrt(totalvariance/ttm)
    /// and then applying the chain rule
    let public impliedvolgradient (a,b,rho,m,sigma) ttm fwd strike  =
        if strike <= 0. || fwd <= 0. || ttm < 0. then None
        else
            let rho = max -1. (min 1. rho)
            let k = log(strike/fwd)-m
            let tmp = impliedvol (a,b,rho,m,sigma) ttm fwd strike |> Option.map (fun x -> 0.5/(x*ttm))
            let tmp1 = totalvariance (0.,1.,rho,m,sigma) fwd strike
            let tmp2 = totalvariance (0.,1.,0.,m,sigma) fwd strike
            if tmp.IsSome && tmp1.IsSome && tmp2.IsSome then
                ( tmp.Value, tmp.Value*(tmp1.Value), tmp.Value*b*k, 
                  -tmp.Value*b*(rho+k/tmp2.Value),tmp.Value*b*sigma/tmp2.Value 
                  ) |> Some
            else None
  
    /// Converts SVI raw parameters to SVI jump-wing parameters
    let public rawtojumpwing (a,b,rho,m,sigma) ttm =
        let helper1 = sqrt(m*m+sigma*sigma)
        let rho = max -1. (min 1. rho)
        let wt = (a+b*(-rho*m+helper1))
        if ttm <= 0. || wt <= 0. || helper1 = 0. then None
        else
            let vt,helper2 = wt/ttm,b/sqrt(wt)
            let phit,pt,ct,vtilde = 0.5*helper2*(rho-m/helper1),helper2*(1.-rho), helper2*(1.+rho),(a+b*sigma*sqrt(1.-rho*rho))/ttm
            Some(vt,phit,pt,ct,vtilde)

    /// Converts SVI jump-wing parametres to SVI raw parameters
    let public jumpwingtoraw (vt,phit,pt,ct,vtilde) ttm =
        if ttm <= 0. || vt <= 0.  || vtilde <= 0. then None
        else
            let helper1 = sqrt(vt*ttm)
            let b = 0.5*helper1*(ct+pt)
            let rho = 1.-pt*helper1/b
            let helper2,beta = sqrt(1.-rho*rho), rho - 2.*phit*helper1/b
            match beta = 0. with
                // m = 0 case, not stated explicitly in Gatheral paper, but easily proven that this is equivalent
                // In Gatheral's paper, the definitions of sigma form a linear system, here the direct solution of that system is used.
                | true  ->  let m,sigma = 0., (vt-vtilde)*ttm/(b*(1.-helper2))
                            let a = vtilde*ttm-b*sigma*helper2
                            Some(a,b,rho,m,sigma)
                // Default case
                | false ->  let sign x = sign(x) |> float
                            if abs(beta) > 1. then None
                            else
                                let alpha = sign(beta)*sqrt(-1. + 1./(beta*beta))    
                                let m = (vt-vtilde)*ttm/(b*(-rho+ sign(alpha)*sqrt(1.+alpha*alpha)-alpha*helper2))
                                let sigma = alpha*m
                                let a = vtilde*ttm-b*sigma*helper2
                                Some(a,b,rho,m,sigma)


    /// Calibrates SVI Raw parameters according to the methodology described in "Quasi-Explicit Calibration of Gatheral's SVI model"
    /// smin is the minimum value allowed for the sigma parameter, mguess is the intial guess for m and sguess is the initial guess for sigma.
    /// the ttm is the time to maturity for the input given, the forward is the forward of the undelrying for the given maturity and points is
    /// an array of tuples where the first element is a weight, the second element is a strike and the third element is the volatility at that strike.
    let public calibrate (checkValid,smin) ttm forward points =
        // If the number of points provided is very small, this doesn't work. In that case, the
        // default paramters below are used. If the initial value calculator gets enough points
        // that initial value is used.
        let DEFAULT_EST_M = 0.
        let DEFAULT_EST_S = 0.1
        
        let smin = max 0. smin
        let third (_,_,x) = x
        let points = points |> Array.map (fun (w,strike,vol) -> (w,log(strike/forward),vol*vol*ttm))
        let maxvar = points |> Array.maxBy third |> third

        /// Checks whether given (transformed zeliade) parameters fall within the boundaries of the region that is potentially arbitrage-free
        /// This function returning true does not imply that parameters are arbitragefree, but if the function returns false, the parameters
        /// are definitely not arbitrage-free.
        let isValid s (a,c,d) =
            not checkValid || ( 0. <= c && c <= 4. * s &&
                                abs(d) <= min c (4.*s-c) &&
                                0. <= a && a <= maxvar)

        /// Converts zeliade parameters into SVI (total variance) parameters
        let zeliadeToSVI (m,s) (a,c,d) = if c = 0. then (a,0.,0.,m,s) else (a,c/s,d/c,m,s)


        /// Provides an initial guess for m and s. Based on stuff from Axel Vogt, but slightly adapted since I encountered
        /// negative sInitial on some test inputs.
        let initialGuess =
            let size = -1 + Array.length points
            if size < 3 then (DEFAULT_EST_M,DEFAULT_EST_S)
            else
                let x,y = points |> Array.map(fun (_,x,v) -> x,v)
                                 |> Array.unzip
                let calcA n = (x.[n]*y.[n+1] - y.[n]*x.[n+1])/(x.[n]-x.[n+1])
                let calcB n = (y.[n+1]-y.[n])/(x.[n+1]-x.[n]) |> fun x -> abs(x)
                
                let aL, bL,aR, bR = calcA 0, -(calcB 0), calcA (size-1), calcB (size-1)
                let recurring = if bR=bL then 0. else bL*(aR-aL)/(bL-bR)      
                 //let aInitial = aL + recurring // Not needed here, but given just for convenience so that some initial guess for all parameters is here 
                let rInitial = if bR=bL then 0. else (bL+bR)/(bR-bL)
                let bInitial = 0.5*(bR-bL)
                let mInitial = recurring/(bInitial * (rInitial-1.))
                let sigmaInitial = -(-(Array.min y) + aL + recurring) / (bInitial * sqrt( abs (1. - rInitial * rInitial) )) |> abs
                mInitial,(max smin sigmaInitial)

        /// For given input m and s, this function calculates the returns a tuple consisting of all the SVI parameters
        /// corresponding to the best "Zeliade solution" found and the total error of the parameters
        let zeliadeSol (m,s) =
            let points = points |> Array.map (fun (w,str,v) -> (abs(w),(str-m)/s,v))
            let sumW = points |> Array.sumBy (fun (wgt,_,_) -> wgt)
            let errorFunc (a,c,d) =
                points |> Array.sumBy (fun (wgt,y,var) -> let err = a + d*y + c * sqrt(y*y+1.) - var
                                                          wgt*err*err)
                       |> fun sum  -> if sumW = 0. then 0. else sum/sumW
                       |> sqrt
            
            // OK the below is a bit of a mess, but it's an efficient mess. Basically a number of numerical expressions occur in the
            // linear system to be solved. The approach taken below is calculate them all as once in a tuple, while folding over
            // the points list. A few helper variables that reoccur in multiple calculations are updated and then the tuple is
            // updated elementwise by just summing.
            let w, wy, wsqr, wv, wy2, wysqr ,wvy, wvsqr =
               points |> Array.fold (fun (acc1,acc2,acc3,acc4,acc5,acc6,acc7,acc8) (wgt,y,var) ->
                                            let wy',wv',sqrh = wgt*y,wgt*var,sqrt(y*y+1.)
                                            (acc1 + wgt, acc2 + wy',acc3 + wgt*sqrh, acc4 + wv',
                                             acc5 + wy'*y,acc6 + wy'*sqrh,acc7+ wy' * var, acc8 + wv' * sqrh)
                                    ) (0.,0.,0.,0.,0.,0.,0.,0.)

            /// Takes an input vector x, the first of which repersents a matrix (in a list) and the second of which represents a vector (as a list)
            /// and solves the linear system defined by them. Since the inputs are intedend to be the equations defining the optimal values for a,d and
            /// c as defined by zeliade, they then get changed order to coincide with the order of SVI arguments in all other functions in this module.
            let solveSystem x = x |> fun (A,b) -> matrix A, vector b
                                  |> fun (A,b) -> A.LU().Solve(b).ToArray()
                                  |> fun arr -> if Array.length arr = 0 then (0.,0.,0.) else (arr.[0],arr.[2],arr.[1]) // Go from [|a;d;c|] to (a,c,d) 

            /// Transforms a tuple of a,c,d parametes i a tuple of all parameters, joined with the error corresponding to those parametes.
            let output tuple = zeliadeToSVI (m,s) tuple, errorFunc tuple

            let globalSol = solveSystem ([[w;wy;wsqr]; [wy; wy2; wysqr]; [wsqr; wysqr; w + wy2]], [wv; wvy; wvsqr])
            if not (isValid s globalSol) then
                /// Finds the solutions for which one of the three parameters is at its bounds
                let facetSols =
                   [[[1.;0.;0.];[wy; wy2; wysqr]; [wsqr; wysqr; w + wy2]],  [0.; wvy; wvsqr];       // system for a = 0
                    [[1.;0.;0.];[wy; wy2; wysqr]; [wsqr; wysqr; w + wy2]],  [maxvar; wvy; wvsqr];   // system for a = maxvar
                    [[w;wy;wsqr]; [0.;-1.;1.]; [wsqr; wysqr; w + wy2]],     [wv; 0.; wvsqr];        // system for d = c
                    [[w;wy;wsqr]; [0.; 1.;1.]; [wsqr; wysqr; w + wy2]],     [wv; 0.; wvsqr];        // system for d = -c
                    [[w;wy;wsqr]; [0.; 1.;1.]; [wsqr; wysqr; w + wy2]],     [wv; 4.*s; wvsqr];      // system part one for |d| <= 4*s-c
                    [[w;wy;wsqr]; [0.;-1.;1.]; [wsqr; wysqr; w + wy2]],     [wv; 4.*s; wvsqr];      // system part two for |d| <= 4*s-c
                    [[w;wy;wsqr]; [wy; wy2; wysqr]; [0.;0.;1.]],            [wv; wvy; 0.];          // system for c = 0
                    [[w;wy;wsqr]; [wy; wy2; wysqr]; [0.;0.;1.]],            [wv; wvy; 4.*s]]        // system for c = 4*s
                    |> List.map solveSystem
                    |> List.filter (isValid s)
                    |> List.map output

                /// Finds the solution for which two of the three parameters are at its bounds. The Zeliade whitepaper speaks about
                /// doing one-dimensinal search to find the minimum, but since fixing 2 out of the three parameters just leads to
                /// a quadratic equation in the third, finding the global minimum and then cutting it off if it falls outside the
                /// valid boundaries, works.
                let cornerOrLineSols =  
                  [[[w;wy;wsqr]; [0.;1.;0.];[0.;0.;1.]],[wv; 0.; 0.];                   // c = 0, implies d = 0, find optimal a
                   [[w;wy;wsqr]; [0.;1.;0.];[0.;0.;1.]],[wv; 0.; 4.*s ];                // c = 4s, implied d = 0, find optimal a
                   [[1.;0.;0.]; [0.;-1.;1.]; [wsqr; wysqr; w + wy2]],[0.;0.;wvsqr];     // a = 0, d = c, find optimal c
                   [[1.;0.;0.]; [0.; 1.;1.]; [wsqr; wysqr; w + wy2]],[0.;0.;wvsqr];     // a = 0, d = -c, find optimal c
                   [[1.;0.;0.]; [0.; 1.;1.]; [wsqr; wysqr; w + wy2]],[0.; 4.*s;wvsqr];  // a = 0, d = 4s-c, find optimal c
                   [[1.;0.;0.]; [0.;-1.;1.]; [wsqr; wysqr; w + wy2]],[0.; 4.*s;wvsqr]]  // a = 0, d = c-4s, find optimal c
                   |> List.map solveSystem 
                   |> List.map (fun (a,c,d) -> let dbound = min c (4.*s-c)
                                               (max 0. (min maxvar a)),(max 0. (min c 4.*s)), max -dbound (min d dbound)) // Cut off at boundary incase we're outside
                   |> List.filter (isValid s) // There will alwasy be at least two solutions here, since c=d=0 will lead to a=0 or a=maxvar and both are valid.
                   |> List.map output

                facetSols @ cornerOrLineSols |> List.minBy snd // Because cornerOrLineSols always contains at least one solution, this array is never empty

            else globalSol |> output


        // TODO: see if this optimization can be improved upon, especially given that the gradient is known explicity.
        let errf (arr : float [])= (arr.[0],arr.[1]) |> zeliadeSol |> snd
        let lb,ub = [|None;Some(smin)|],[|None;None|]
        minimize errf (initialGuess |> fun (m,s) -> [|m;s|]) lb ub
            |> fun arr -> arr.[0],arr.[1]
            |> zeliadeSol 
        //zeliadeSol initialGuess //These are just the parameters for a smart intial guess. Typically perform not too shabby, but optimization can still lead to good improvements.
                                  //Just showing it here because it could be useful at some point.