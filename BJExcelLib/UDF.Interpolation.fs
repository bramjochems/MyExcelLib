namespace BJExcelLib.UDF

/// Contains UDFs for interpolation functions. Two types of functions exists. One are UDFS whose excel name is of the form Interpolate.XXX. Those
/// UDFs are UDFs that directly do interpolation. For efficiency reasons, when interpolating multiple data points, array formulae on the new x-values
/// can be used. The other are UDFs of the form BuildInterpolator.XXX. These UDFs create a function in memory which can later be used for interpolation
/// by calling Interpolate.FromObject. The idea here is that array formulae are never needed. These functions are useful if the input data is relatively
/// fixed but the value to interpolate at can change one-by-one (in which case recalculating all of them is inefficient).
module public Interpolation =
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.ExcelDna.Cache
    open BJExcelLib.Math.Interpolation
    open BJExcelLib.Math.Interpolation2D
    open BJExcelLib.Util.Extensions

    let private INTERPOLATOR_TAG = "BJ1DInterpolator"
    let private INTERPOLATOR2D_TAG = "BJ2DInterpolator"

    let private validateArray = array1DAsArray validateFloat

    /// Helper function for most of the interpolation calculation functions
    let private ipcalc f xold yold xnew =
        let xold = validateArray xold
        let yold = validateArray yold
        let ipfunc =
            if Array.length xold = Array.length yold then
                Array.zip xold yold
                    |> Array.filter(fun (x,y) -> x.IsSome && y.IsSome)
                    |> Array.map (fun (x,y) -> x.Value,y.Value)
                    |> f
            else None
        match ipfunc with
            | None      -> xnew |> Array.map (fun _  -> None)
            | Some(f)   -> xnew |> validateArray
                                |> Array.map (Option.map f)
        |> optionArrayToExcelColumn
   
    /// Helper function for building an interpolator object for most of the interpolation functions
    let private ipbuild f xold yold =
        let xold = validateArray xold
        let yold = validateArray yold
        if Array.length xold = Array.length yold then
            Array.zip xold yold
                |> Array.filter(fun (x,y) -> x.IsSome && y.IsSome)
                |> Array.map (fun (x,y) -> x.Value,y.Value)
                |> f
        else None
        |> fun x -> match x with
                        | None    -> None |> valueToExcel
                        | Some(f) -> register INTERPOLATOR_TAG f
     
    /// Interpolation function that interpolates using a given object in the cache. Interpolates on only a single value because the entire purpose of building an object
    /// in the cache is that you avoid array formulas and trigger only a single calculation if a single value changes.
    [<ExcelFunction(Name="Interpolate.FromObject", Description="Interpolation from a previously constructed interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfipfromobj ([<ExcelArgument(Name="Interpolator",Description="Name of the interpolator object")>] name : obj,
                             [<ExcelArgument(Name="Xnew",Description = "New x-value to interpolate at")>] xnew : obj) =
        validateString name
            |> fun s -> if s.IsSome && (s.Value.Contains(INTERPOLATOR_TAG)) then (Option.bind lookup s) else None
            |> fun s -> match s with
                        | None      -> None
                        | Some(f)   -> xnew |> validateFloat |> Option.map ( f :?> (float->float) )
            |> valueToExcel          
    
    /// Performs piecewise constant interpolation on input data                            
    [<ExcelFunction(Name="Interpolate.PiecewiseConstant", Description="Piecewise constant 1D Interpolation",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipconstant ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                              [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                              [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorPiecewiseConstant xold yold xnew

    /// Builds a piecewise contant interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.PiecewiseConstant", Description="Builds piecewise constant 1D Interpolator",Category="BJExcelLib.Interpolation")>]
    let public udfbldconstant([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                              [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorPiecewiseConstant xold yold

    /// Performs piecewise linear interpolation on input data                            
    [<ExcelFunction(Name="Interpolate.PiecewiseLinear", Description="Piecewise constant 1D Interpolation",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfiplinear ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorPiecewiseLinear xold yold xnew

    /// Builds a piecewise contant interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.PiecewiseLinear", Description="Builds piecewise constant 1D Interpolator",Category="BJExcelLib.Interpolation")>]
    let public udfbldlinear ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                             [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorPiecewiseLinear xold yold

    /// Performs piecewise Loglinear interpolation on input data                            
    [<ExcelFunction(Name="Interpolate.PiecewiseLoglinear", Description="Piecewise loglinear 1D Interpolation",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipLoglinear  ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                                [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorPiecewiseLoglinear xold yold xnew

    /// Builds a piecewise contant interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.PiecewiseLoglinear", Description="Builds a piecewise loglinear interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfbldLoglinear ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorPiecewiseLoglinear xold yold

    /// Performs interpolation on the input data using an akima spline                          
    [<ExcelFunction(Name="Interpolate.AkimaSpline", Description="1D Interpolation using an Akima spline",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipakima  ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorAkimaSpline xold yold xnew

    /// Builds a akima spline interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.AkimaSpline", Description="Builds an Akima spline interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfbldakima ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorAkimaSpline xold yold

    /// Performs interpolation on the input data using a natural spline                          
    [<ExcelFunction(Name="Interpolate.NaturalSpline", Description="1D Interpolation using a natural spline",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipnatural([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorNaturalSpline xold yold xnew

    /// Builds a natural spline interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.NaturalSpline", Description="Builds a natural spline interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfbldnatural   ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorNaturalSpline xold yold

    /// Performs interpolation on the input data using a natural spline                          
    [<ExcelFunction(Name="Interpolate.ClampedSpline", Description="1D Interpolation using a clamped spline",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipclamped([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="dLeft",Description = "The derivative at the leftmost x-point")>] left : obj,
                            [<ExcelArgument(Name="dRight",Description = "The derivative at the rightmost x-point")>] right : obj,
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        let l = validateFloat left |> Option.getOrElse 0.
        let r = validateFloat right |> Option.getOrElse 0.
        ipcalc (interpolatorClampedSpline l r) xold yold xnew

    /// Builds a natural spline interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.ClampedSpline", Description="Builds a clamped spline interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfbldclamped   ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                                [<ExcelArgument(Name="dLeft",Description = "The derivative at the leftmost x-point")>] left : obj,
                                [<ExcelArgument(Name="dRight",Description = "The derivative at the rightmost x-point")>] right : obj) =
        let l = validateFloat left |> Option.getOrElse 0.
        let r = validateFloat right |> Option.getOrElse 0.
        ipbuild (interpolatorClampedSpline l r) xold yold

    /// Performs linear gitting on input data                            
    [<ExcelFunction(Name="Interpolate.FitLinear", Description="Least-square linear fitting (1D)",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udffitlinear([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc fitLinear xold yold xnew

    /// Builds a piecewise contant interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.FitLinear", Description="Builds linear fitting object",Category="BJExcelLib.Interpolation")>]
    let public udfbldlinfit([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild fitLinear xold yold

    /// Performs polynomial fitting                          
    [<ExcelFunction(Name="Interpolate.FitPolynomial", Description="Least-square polynomial fitting (1D)",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udffitpoly  ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Degree",Description = "Maximum degree of the fitting polynomial. Optional, default = 1")>] degree : obj,
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        let degree = degree |> validateInt |> Option.getOrElse 1
        ipcalc (fitPolynomial degree) xold yold xnew

    /// Builds a piecewise contant interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.FitPolynomial", Description="Builds polynomial fitting object",Category="BJExcelLib.Interpolation")>]
    let public udfbldpolyft([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Degree",Description = "Maximum degree of the fitting polynomial")>] degree : obj) =
        let degree = degree |> validateInt |> Option.getOrElse 1
        ipbuild (fitPolynomial degree) xold yold

    /// Performs interpolation on the input data using a Fritsch-Butland spline                          
    [<ExcelFunction(Name="Interpolate.FritschButlandSpline", Description="1D Interpolation using a Fritsch-Butland spline (which is a type of monotone spline)",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfipfbsplin([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        ipcalc interpolatorFritschButlandSpline xold yold xnew

    /// Builds a Fritsch-Butland spline interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.FritschButlandSpline", Description="Builds a Fritsch-Butland spline interpolator object (which is a type of monotone spline)",Category="BJExcelLib.Interpolation")>]
    let public udfbldfbspline  ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj []) =
       ipbuild interpolatorFritschButlandSpline xold yold

    /// Performs interpolation on the input data using a hermitel spline                          
    [<ExcelFunction(Name="Interpolate.HermiteSpline", Description="1D Interpolation using a Hermite spline",Category="BJExcelLib.Interpolation", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public udfiphermite([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                            [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                            [<ExcelArgument(Name="dYold",Description = "Derivative values at the x-points")>] y'old : obj [],
                            [<ExcelArgument(Name="Xnew",Description = "New x-values to interpolate at")>] xnew : obj []) =
        let xold = validateArray xold
        let yold = validateArray yold
        let y'old = validateArray y'old
        let ipfunc =
            if Array.length xold = Array.length yold && Array.length xold = Array.length y'old then
                Array.zip3 xold yold y'old
                    |> Array.filter(fun (x,y,y') -> x.IsSome && y.IsSome && y'.IsSome)
                    |> Array.map (fun (x,y,y') -> x.Value,y.Value, y'.Value)
                    |> interpolatorHermiteSpline
            else None
        match ipfunc with
            | None      -> xnew |> Array.map (fun _  -> None)
            | Some(f)   -> xnew |> validateArray
                                |> Array.map (Option.map f)
        |> optionArrayToExcelColumn

    /// Builds a hermite spline interpolator object in the cache
    // functions that create an object in the cache cannot be marked threadsafe and exception safe
    [<ExcelFunction(Name="BuildInterpolator.HermiteSpline", Description="Builds a Hermite spline interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfbldhermite  ([<ExcelArgument(Name="Xold",Description = "Given input x-values")>] xold : obj [],
                                [<ExcelArgument(Name="Yold",Description = "Given input y-values")>] yold : obj [],
                                [<ExcelArgument(Name="dYold",Description = "Derivative values at the x-points")>] y'old : obj []) =
        let xold = validateArray xold
        let yold = validateArray yold
        let y'old = validateArray y'old
        if Array.length xold = Array.length yold && Array.length xold = Array.length y'old then
            Array.zip3 xold yold y'old
                |> Array.filter(fun (x,y,y') -> x.IsSome && y.IsSome && y'.IsSome)
                |> Array.map (fun (x,y,y') -> x.Value,y.Value, y'.Value)
                |> interpolatorHermiteSpline
        else None
        |> fun x -> match x with
                        | None    -> None |> valueToExcel
                        | Some(f) -> register INTERPOLATOR_TAG f

    
    let private buildip2D xkeys ykeys table ipFunc =
        let xkeys = validateArray xkeys
        let ykeys = validateArray ykeys
        let table = array2DAsArray true validateFloat table
        build2DInterpolation ipFunc xkeys ykeys table
    
    let private registerip2D xkeys ykeys table ipFunc =
        buildip2D xkeys ykeys table ipFunc
        |> register INTERPOLATOR2D_TAG

    let private ip2D xkeys ykeys table xnew ynew ipFunc=
        let ipfunc = buildip2D xkeys ykeys table ipFunc
        let xnew = validateArray xnew
        let ynew = validateArray ynew
        let res = Array2D.init  (Array.length xnew)
                                (Array.length ynew)
                                (fun r c -> if xnew.[r].IsSome && ynew.[c].IsSome then ipfunc (xnew.[r].Value) (ynew.[c].Value) else None)
        res |> Array2D.map valueToExcel 

    [<ExcelFunction(Name="Interpolate2D.Bilinear", Description="Bilinear interpolation",Category="BJExcelLib.Interpolation")>]
    let public udfipblinear([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,],
                            [<ExcelArgument(Name="Xnew",Description = "x-value to interpolate on")>] xnew : obj [],
                            [<ExcelArgument(Name="Ynew",Description = "y-value to interpolate on")>] ynew : obj []) =
        ip2D xkeys ykeys table xnew ynew interpolatorPiecewiseLinear

    [<ExcelFunction(Name="BuildInterpolator2D.Bilinear", Description="Builds a 2D bilinear nterpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfblblinear([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,]) =
        registerip2D xkeys ykeys table interpolatorPiecewiseLinear

    [<ExcelFunction(Name="Interpolate2D.Bicubic", Description="Bicubic spline interpolation using natural splines",Category="BJExcelLib.Interpolation")>]
    let public udfipbicubic([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,],
                            [<ExcelArgument(Name="Xnew",Description = "x-value to interpolate on")>] xnew : obj [],
                            [<ExcelArgument(Name="Ynew",Description = "y-value to interpolate on")>] ynew : obj []) =
        ip2D xkeys ykeys table xnew ynew interpolatorNaturalSpline

    [<ExcelFunction(Name="BuildInterpolator2D.Bicubic", Description="Builds a 2D bicubic interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfblbicubic([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,]) =
        registerip2D xkeys ykeys table interpolatorNaturalSpline

    [<ExcelFunction(Name="Interpolate2D.BicubicAkima", Description="Bicubic spline interpolation using akima splines",Category="BJExcelLib.Interpolation")>]
    let public udfipbcubicA([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,],
                            [<ExcelArgument(Name="Xnew",Description = "x-value to interpolate on")>] xnew : obj [],
                            [<ExcelArgument(Name="Ynew",Description = "y-value to interpolate on")>] ynew : obj []) =
        ip2D xkeys ykeys table xnew ynew interpolatorAkimaSpline

    [<ExcelFunction(Name="BuildInterpolator2D.BicubicAkima", Description="Builds a 2D bicubic akima spline interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfblbicubiA([<ExcelArgument(Name="Row indices",Description = "Keys on the x-values")>] xkeys : obj [],
                            [<ExcelArgument(Name="Column indices",Description = "Keys on the y-values")>] ykeys : obj [],
                            [<ExcelArgument(Name="Table",Description = "Function values for x,y points")>] table: obj [,]) =
        registerip2D xkeys ykeys table interpolatorAkimaSpline

    /// Interpolation function that interpolates using a given object in the cache. Interpolates on only a single value because the entire purpose of building an object
    /// in the cache is that you avoid array formulas and trigger only a single calculation if a single value changes.
    [<ExcelFunction(Name="Interpolate2D.FromObject", Description="2D Interpolation from a previously constructed interpolator object",Category="BJExcelLib.Interpolation")>]
    let public udfip2frmobj ([<ExcelArgument(Name="Interpolator",Description="Name of the interpolator object")>] name : obj,
                             [<ExcelArgument(Name="Xnew",Description = "New x-value to interpolate at")>] xnew : obj,
                             [<ExcelArgument(Name="Ynew",Description = "New y-value to interpolate at")>] ynew : obj) =
        validateString name
            |> fun s -> if s.IsSome && (s.Value.Contains(INTERPOLATOR2D_TAG)) then (Option.bind lookup s) else None
            |> fun s -> match s with
                        | None      -> None
                        | Some(f)   -> let xnew = xnew |> validateFloat
                                       let ynew = ynew |> validateFloat
                                       if xnew.IsSome && ynew.IsSome then
                                           (f :?> (float->float-> float option)) (xnew.Value) (ynew.Value)
                                       else None

            |> valueToExcel 
