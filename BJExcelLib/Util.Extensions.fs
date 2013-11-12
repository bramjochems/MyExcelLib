namespace BJExcelLib.Util

// Defines a number of useful extension methods
module public Extensions =
    
    let private countfunc f = fun x -> if f x then 1. else 0.

    // Extension methods on list
    type public Microsoft.FSharp.Collections.List<'a> with
        /// Returns the number of elements in the list that satisfy the condition defined by f
        static member countBy (f : 'a -> bool) x = List.sumBy (countfunc f) x
        
        /// Folds partway through a list returning the accumulator and the remaining tail
        static member foldTo i f accu list =
            let rec folder j g acc l =
                match j, l with | 0,t -> acc,t
                                | i, h::t -> folder (i-1) f (f accu h) t
                                | _, [] -> accu, []
            folder i f accu list

        /// Splits a list into two sublists at the given index
        static member chop i list =
            let front,back = List<'a>.foldTo i (fun t h -> h :: t) [] list
            (List.rev front), back

    // Extension methods on System.DateTime
    type public System.DateTime with
        /// Override for the System.DateTime.tryParse that is more idiotmatic for F#
        static member tryParse (s : string) =
            let ok, res = System.DateTime.TryParse(s)
            if ok then Some(res) else None
 
    /// Extensions to the Array module
    module public Array =
        /// Returns the number of elements int he array that satisfy the condition defined by f
        let countBy (f : 'a -> bool) x = Array.sumBy (countfunc f) x
            
    /// Extension methods on Array2D    
    module public Array2D =
        // Transposes a 2D array. Very simple at the moment with no bound checking or memory optimizations.
        let transpose (data : 'a [,]) =
            let l1 = data.GetLength(0)
            let l2 = data.GetLength(1)
            Array2D.init l2 l1 (fun x y -> data.[y,x])
                



        

                       