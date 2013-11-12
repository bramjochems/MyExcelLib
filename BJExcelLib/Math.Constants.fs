namespace BJExcelLib.Math

/// Provides and wraps some mathematical constants
[<AutoOpen>]
module Constants =

    /// Wraps pi
    let public pi = System.Math.PI

    /// Provides the machine_epsilon (smallest float x such that 1+x != x)
    let public machine_epsilon =
        let rec helper input =
            let x = input/2.
            match 1.+ x with | 1.    -> input
                             | _     -> helper x
        helper 1.
        