namespace BJExcelLib.Math

open BJExcelLib.Util.Extensions

/// Contains some optimization functions. These are homebrewn until an implementation in MathNet.Numerics is available. The reason
/// for that is that I would like to avoid using multiple libraries for different math functionality.
module Optimization =

    let private DIFFERENTATION_STEPSIZE = machine_epsilon |> sqrt

    // TODO: see if I can get rid of using alglib here, would like to avoid the dependency altogether if possible
    let public minimize f initial lb ub =
        let lb = lb |> Array.map (Option.getOrElse System.Double.NegativeInfinity)
        let ub = ub |> Array.map (Option.getOrElse System.Double.PositiveInfinity)

        let alglibfvec = new alglib.ndimensional_fvec( fun x func obj ->  func.[0] <- f x  )
        let epsg,epsf,epsx,maxits = 0.000000001, 0. ,0.,0
        
        let sol = initial
        let state = alglib.minlmcreatev(Array.length sol, sol, 0.0001)
        do alglib.minlmsetcond(state, epsg, epsf, epsx, maxits)
        do alglib.minlmsetbc(state,lb,ub)
        do alglib.minlmoptimize(state, alglibfvec, null, null)
        let final, rep = alglib.minlmresults(state)
        final