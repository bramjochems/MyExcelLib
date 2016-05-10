namespace BJExcelLib.Math

open BJExcelLib.Util.Extensions

/// Contains functions for general interpolation and curve fitting.
module Interpolation =
    open MathNet.Numerics.Fit

    /// Helper function that peforms some typical preprocessing needed for a lot of interpolation functions
    let private processTupleArray xydata =
        xydata  |> Seq.distinctBy fst
                |> Seq.sortBy fst
                |> Array.ofSeq
                |> Array.unzip
             
    /// Piecewise constant interpolation of given (x,y) input tuples
    let public interpolatorPiecewiseConstant (xydata : (float*float) [])=
        if Array.length xydata = 0 then None
        else
            let xdata, ydata = processTupleArray xydata

            //xdata,ydata go into the closure. As a result, they are evaluated only once,but for big arrays
            // this might lead to a lot of memory consumption? Probably not an issue for my intended usage.                          
            Some( fun x -> let ix = xdata   |> Array.tryFindIndex (fun el -> el > x)
                                            |> Option.getOrElse (Array.length xdata)
                                            |> max -1
                           ydata.[ix+1])

    /// Piecewise linear interpolation of given (x,y) input data
    let public interpolatorPiecewiseLinear (xydata : (float*float) []) =
        if Array.length xydata < 2 then interpolatorPiecewiseConstant xydata
        else
            let xdata, ydata = processTupleArray xydata
            Some (fun x ->  let ix = xdata  |> Array.tryFindIndex(fun el -> el > x)
                                            |> Option.map (fun z-> (z-1,z))
                                            |> Option.getOrElse (-2 + Array.length xdata, -1 + Array.length xdata)
                                            |> fun (u,v) -> if u = -1 then (0,1) else (u,v)
                            let x1,x2, y1, y2 = xdata.[fst ix], xdata.[snd ix], ydata.[fst ix], ydata.[snd ix]
                            y1 + (y2-y1)/(x2-x1)*(x-x1))

    /// Piecewise linear interpolation of given (x,y) input data
    let public interpolatorPiecewiseLoglinear (xydata : (float*float) []) =
        if Array.length xydata < 2 then interpolatorPiecewiseConstant xydata
        else
            let xdata, ydata = processTupleArray xydata
            Some (fun x ->  let ix = xdata  |> Array.tryFindIndex(fun el -> el > x)
                                            |> Option.map (fun z-> (z-1,z))
                                            |> Option.getOrElse (-2 + Array.length xdata, -1 + Array.length xdata)
                                            |> fun (u,v) -> if u = -1 then (0,1) else (u,v)
                            let x1,x2, y1, y2 = xdata.[fst ix], xdata.[snd ix], ydata.[fst ix], ydata.[snd ix]
                            y1*((y2/y1)**((x-x1)/(x2-x1))))

    /// Creates an akima spline that interpolators the given input (x,y) data
    let public interpolatorAkimaSpline xydata =
        if Array.length xydata < 5 then None
        else
            let xdata,ydata = processTupleArray xydata
            let ip = MathNet.Numerics.Interpolation.Algorithms.AkimaSplineInterpolation(xdata,ydata)
            Some (fun x -> ip.Interpolate x)

    /// Creates a natural cubic spline that interpolators given input (x,y) data
    let public interpolatorNaturalSpline xydata =
        if Array.length xydata <= 1 then interpolatorPiecewiseConstant xydata
        else
            let xdata,ydata = processTupleArray xydata
            let ip = MathNet.Numerics.Interpolation.Algorithms.CubicSplineInterpolation(xdata,ydata)
            Some (fun x -> ip.Interpolate x)

    /// Creates a clamped cubic spline that inperpolates given input (x,y) data and has specified derivatives at it's endpoints.
    let public interpolatorClampedSpline leftD rightD xydata =
        if Array.length xydata <= 1 then interpolatorPiecewiseConstant xydata
        else
            let xdata,ydata = processTupleArray xydata
            try // Not sure if this can ever fail for any inputs, try catch block isn't expensive here as long as no error occurs.
                let ip = MathNet.Numerics.Interpolation.Algorithms.CubicSplineInterpolation(xdata,ydata,MathNet.Numerics.Interpolation.SplineBoundaryCondition.FirstDerivative,leftD,
                                                                                                        MathNet.Numerics.Interpolation.SplineBoundaryCondition.FirstDerivative,rightD)                                                                                      
                Some (fun x -> ip.Interpolate x)
            with _ -> None

    /// Creates a hermite spline that interpolators given input (x,y,y') data
    let public interpolatorHermiteSpline xyy'data =
        if Array.length xyy'data = 0 then None
        elif Array.length xyy'data = 1 then
            let x1, y1, d1 = xyy'data.[0]
            Some(fun x -> y1 + d1*(x-x1))
        else
            let xdata, ydata, y'data = xyy'data |> Seq.distinctBy (fun (a,_,_) -> a)
                                                |> Seq.sortBy (fun (a,_,_) -> a)
                                                |> Array.ofSeq
                                                |> Array.unzip3
            try
                let ip = MathNet.Numerics.Interpolation.Algorithms.CubicHermiteSplineInterpolation(xdata,ydata,y'data)
                Some(fun x -> ip.Interpolate x)
            with _ -> None


    /// Creates a function that provides a linear least-square fit through given (x,y) input tuples
    let public fitLinear xydata =
        if Array.length xydata <= 1 then interpolatorPiecewiseConstant xydata
        else
            let xdata, ydata = xydata |> Array.unzip
            Some(lineF xdata ydata)

    /// Creates a function that provides a polynomial least-square fit with a degree that is at most equal to "degree" thorugh given (x,y) input tuples
    let public fitPolynomial degree xydata =
        if degree >= 0 && Array.length xydata > degree then
            let xdata, ydata = xydata |> Array.unzip
            Some(polynomialF degree xdata ydata)
        else None

    

    /// Calculates derivative values using the Arithmetic Mean Method (Wang 2006). Input are xdata and ydata arrays that are assumed to have been preprocessed to be sorted by x-value.
    let private AMMDerivatives xdata ydata =
        let n = Array.length xdata
        if n <> Array.length ydata then raise (new System.ArgumentOutOfRangeException("Arrays of x-values and y-values should have same size"))
        match n with
            | 0 -> [||]
            | 1 -> [|0.|]
            | 2 -> [|(ydata.[1]-ydata.[0])/(xdata.[1]-xdata.[0])|]
            | _ -> let h,D = Seq.zip xdata ydata    |> Seq.windowed 2
                                                    |> Seq.map (fun [| (x1,y1); (x2,y2) |] -> x2-x1,(y2-y1)/(x2-x1))
                                                    |> Seq.toArray
                                                    |> Array.unzip
                   [| for i in 0 .. n-1 do
                            if i = 0 then yield D.[0] + (D.[0]-D.[1])*h.[0]/(h.[0]+h.[1])
                            elif i = n-1 then yield D.[i-1] + (D.[i-1]-D.[i-2])*h.[i-1]/(h.[i-1] + h.[i-2])
                            else yield (h.[i]*D.[i-1] + h.[i-1]*D.[i])/(h.[i-1]+h.[i]) |]

    /// Calculates derivative values using the Kruger method which prevents overshoot problems.
    let private krugerDerivatives xdata ydata =
        let n = Array.length xdata
        if n <> Array.length ydata then raise (new System.ArgumentOutOfRangeException("Arrays of x-values and y-values should have same size"))
        match n with
            | 0 -> [||]
            | 1 -> [|0.|]
            | 2 -> [|(ydata.[1]-ydata.[0])/(xdata.[1]-xdata.[0])|]
            | _ -> let h,D = Seq.zip xdata ydata    |> Seq.windowed 2
                                                    |> Seq.map (fun arr ->  let (x1,y1) = arr.[0]
                                                                            let (x2,y2) = arr.[1]
                                                                            x2-x1,(y2-y1)/(x2-x1))
                                                    |> Seq.toArray
                                                    |> Array.unzip
                   let res = [| for i in 0 .. n-1 do
                                    if i = 0 || i=n-1 then yield 0.
                                    else yield 2./(1./D.[i-1] + 1./D.[i]) |]
                   res.[0] <- 1.5*D.[0] - 0.5*res.[1]
                   res.[n-1] <- 1.5*D.[n-2] - 0.5*res.[n-2]
                   res                                          

    /// Calculates derivative values using the Fritsch-Butland method which ensures local monoticity
    let private FBDerivatives xdata ydata =
        let n = Array.length xdata
        if n <> Array.length ydata then raise (new System.ArgumentOutOfRangeException("Arrays of x-values and y-values should have same size"))
        match n with
            | 0 -> [||]
            | 1 -> [|0.|]
            | 2 -> [|(ydata.[1]-ydata.[0])/(xdata.[1]-xdata.[0])|]
            | _ -> let h,D = Seq.zip xdata ydata    |> Seq.windowed 2
                                                    |> Seq.map (fun arr ->  let (x1,y1) = arr.[0]
                                                                            let (x2,y2) = arr.[1]
                                                                            x2-x1,(y2-y1)/(x2-x1))
                                                    |> Seq.toArray
                                                    |> Array.unzip
                   let res = [| for i in 0 .. n-1 do
                                    if i = 0 || i=n-1 then yield 0.
                                    else
                                        let dmin = min D.[i] D.[i-1] 
                                        let dmax = max D.[i] D.[i-1]
                                        if dmax = 0. && dmin = 0. then yield 0. else yield 3.*dmin*dmax/(dmax+2.*dmin) |]
                   res.[0]      <- ((2.*h.[0]+h.[1])*D.[0] - h.[0]*D.[1])/(h.[0]+h.[1])
                   res.[n-1]    <- ((2.*h.[n-2] + h.[n-3])*D.[n-2] - h.[n-2]*D.[n-3])/(h.[n-2]+h.[n-3])
                   res

    /// Calculates derivative values using a geometric mean approach
    let private GMDerivatives xdata ydata =
        let n = Array.length xdata
        if n <> Array.length ydata then raise (new System.ArgumentOutOfRangeException("Arrays of x-values and y-values should have same size"))
        match n with
            | 0 ->  [||]
            | 1 ->  [|0.|]
            | 2 ->  [|(ydata.[1]-ydata.[0])/(xdata.[1]-xdata.[0])|]
            | _ ->  let h,D = Seq.zip xdata ydata   |> Seq.windowed 2
                                                    |> Seq.map (fun arr ->  let (x1,y1) = arr.[0]
                                                                            let (x2,y2) = arr.[1]
                                                                            x2-x1,(y2-y1)/(x2-x1))
                                                    |> Seq.toArray
                                                    |> Array.unzip
                    let D31 = (ydata.[2]-ydata.[0])/(xdata.[2]-xdata.[0])
                    let Dnn = (ydata.[n-1]-ydata.[n-3])/(xdata.[n-1]-xdata.[n-3])
                    [| for i in 0 .. n-1 do
                            if i = 0 then
                                if D31 = 0. || D.[0] = 0. then yield 0.
                                else 
                                    let frac = h.[0]/h.[1]
                                    yield (D.[0]**(1.+frac))*(D31**(-frac))
                            elif i = n-1 then
                                if Dnn = 0. || D.[n-2] = 0. then yield 0.
                                else
                                    let frac = h.[n-2]/h.[n-3]
                                    yield (D.[n-2]**(1.+frac))*(Dnn**(-frac))
                            else
                                if D.[i] = 0. || D.[i-1] = 0. then yield 0.
                                else
                                    yield (D.[i-1]**(h.[i]/(h.[i-1]+h.[i])))*(D.[i]**(h.[i-1]/(h.[i-1]+h.[i]))) |]

    /// Creates a Fritsch-Butland spline, which I think (todo: check) is monotone
    let public interpolatorFritschButlandSpline xydata =
        if Array.length xydata <= 2 then interpolatorNaturalSpline xydata
        else
            let xdata,ydata = processTupleArray xydata
            let y'data = FBDerivatives xdata ydata
            Array.zip3 xdata ydata y'data |> interpolatorHermiteSpline
                             

    // todo : dougherty/hyman modification of arbitrary derivative values to produce monotone splines
        