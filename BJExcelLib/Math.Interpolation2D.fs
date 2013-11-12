namespace BJExcelLib.Math

/// Contains functions for general interpolation and curve fitting.
module Interpolation2D =
    open BJExcelLib.Math.Interpolation

    let flatten (A:'a[,]) = A |> Seq.cast<'a>
    
    let getColumn c (A:_[,]) = flatten A.[*,c..c] |> Seq.toArray

    let getRow r (A:_[,]) = flatten A.[r..r,*] |> Seq.toArray
    let public build2DInterpolation (ipfunc : (float*float) [] -> (float -> float) option) (xkeys : float option []) (ykeys : float option []) (table : float option [,])=
        let nrows = Array.length xkeys
        let ncols = Array.length ykeys
        fun x y ->
            /// For every row of the table of x values, interpolate the current row at the value y
            let xtable =
                [| for r in 0 .. nrows - 1 do
                    if xkeys.[r].IsSome then
                       let rowdata = table |> getRow r
                       let rowIp  = Array.zip ykeys rowdata |> Array.filter (fun (x,y) -> x.IsSome && y.IsSome)
                                                            |> Array.map (fun (x,y) -> x.Value, y.Value)
                                                            |> ipfunc
                       yield rowIp|]
                /// Create a table of the form xkey,interpolated y value and interpolate this table
                |> Array.mapi (fun r f -> xkeys.[r], if f.IsSome then Some(f.Value y) else None)
                |> Array.filter (fun (r,rowIp) -> r.IsSome && rowIp.IsSome)
                |> Array.map (fun (r,rowIp) -> r.Value,rowIp.Value)

            ipfunc xtable |> fun z -> match z with | Some(f) -> (f x) |> Some
                                                   | None    -> None
            
            