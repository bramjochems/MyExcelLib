namespace BJExcelLib.Util

[<AutoOpen>]
module Misc =

    let inline isNull (x:^T when ^T : not struct) = obj.ReferenceEquals (x, null)

 