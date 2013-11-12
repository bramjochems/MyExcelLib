namespace BJExcelLib.ExcelDna

/// This module provides IO functions for UDFs created by ExcelDNA.
/// Because of a combination of the type inference in F# and performance reasons, all ExcelDNA function inputs in this project have input type "obj". This module contains functionality
/// to cast those variables to other types and conversly convert other types back to obj's (or obj []'s) to return to excel. Next to this, there is some functionality to deal with returning
/// option types to excel.
module public IO =

    open BJExcelLib.Util.Extensions
    open ExcelDna.Integration

    let (|ExcelError|_|) x =
        if x.GetType() = typeof<ExcelError> then Some() else None
    
    let (|ExcelEmpty|_|) x =
        if x.GetType() = typeof<ExcelEmpty> then Some() else None
   
    let (|ExcelMissing|_|) x =    
        if x.GetType() = typeof<ExcelMissing> then Some() else None

    /// Takes an obj value and check if it's an valid argument coming from excel or not. Returns an obj option.
    let private wrapErrorValues x  =
        match x with
            | ExcelError | ExcelEmpty | ExcelMissing    ->  None
            | _                                         ->  Some(x)

    /// wraps primitive values + strings and date times into a Some(value), everything else becomes None
    let private wrapprimitive x = if x.GetType().IsPrimitive then Some(x)
                                  elif x.GetType() = typeof<string> then Some(x) // Strings are not primitive types in .NET but for my purpose here I'll consider them to be... 
                                  elif x.GetType() = typeof<System.DateTime> then Some(x) // DateTimes are not primitive types in .NET but for my purpose here I'll consider them to be
                                  else None

    /// Validates an input to a specific type if possible. If not, None is returned.
    let private validateType<'T> input =
      try Some(System.Convert.ChangeType(input, typeof<'T>) :?> 'T)
      with _ -> None //This error catching mechanism is slow, but shouldn't occur to frequently if the wrapErrorValues function is used to catch wrong input first

    /// Validates an input to a specific given type if possible. If not, None is returned.
    let private validate<'T> input =
        input   |> wrapErrorValues
                |> Option.bind validateType<'T>

    /// Takes an obj option and tries to cast it into Some(string). If not possible, None is returned
    let public validateString input = validate<string> input

    /// Takes an obj option and tries to cast it into Some(float). If not possible, None is returned
    let public validateFloat input = validate<float> input

    /// Takes an obj option and tries to cast it into Some(int). If not possible, None is returned
    let public validateInt input = validate<int> input

    /// Takes an obj option and tries to cast it into Some(bool). If not possible, None is returned. Numeric types are mapped to true if they are > 0
    let public validateBool input =
        input   |> validate<bool>
                |> fun x -> if x.IsSome then x
                            else input  |> validateFloat
                                        |> Option.map (fun x -> x > 0.) //todo : test whether the same issue occurs as with validatedate, i.e. that booleans are passed as numeric types. In that case, it's probably better to first check on numeric types for performance reasons.

    /// Takes an obj option and tries to cast it into Some(DateTime). If not possible, None is returned. Numeric types are mapped to true if they are > 0                  
    let public validateDate input =
        input   |> validate<float> |> Option.map (fun x -> System.DateTime.FromOADate(x)) // First check for float rather than date because that's how obj type's typically come through.
                                                                                          // Makes a huge difference on performance, vs try-catching System.DateTime first, because catching
                                                                                          // errors is so slow.
                |> fun x -> if x.IsSome then x else input  |> validate<System.DateTime>
                |> fun x -> if x.IsSome then x else input  |> validate<string> |> Option.bind (fun x -> System.DateTime.tryParse(x))

    /// transfoms an option type to an UDF return value that excel can handle
    /// None values get transformed to an excelerror.
    let public valueToExcel value =
        value |> Option.map box
              |> FSharpx.Option.getOrElse (ExcelError.ExcelErrorNA :> obj)

    /// Transforms a 1D array into a sequence
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         arr is  the array to transform
    let array1DAsSequence validationfunction =
        Seq.map (wrapErrorValues >> Option.bind validationfunction)

    /// Transforms a 1D array into a sequence
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         defaultvalue is the default value to replace missing values with
    ///         arr is the array to transform
    let array1DAsSequenceWithDefault validationfunction defaultvalue =
        Seq.map (wrapErrorValues >> Option.bind validationfunction >> (FSharpx.Option.getOrElse defaultvalue))

    /// Transforms a 1D array into a List
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         arr = the array to transform
    let array1DAsList validationfunction arr =
        arr |> array1DAsSequence validationfunction 
            |> Seq.toList

    /// Transforms a 1D array into a List
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         defaultvalue is the default value to replace missing values with
    ///         arr is the array to transform
    let array1DAsListWithDefault validationfunction defaultvalue arr =
        arr |> array1DAsSequenceWithDefault validationfunction defaultvalue
            |> Seq.toList

    /// Transforms a 1D array into an Array with typed (casted) elements
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         arr = the array to transform
    let public array1DAsArray validationfunction =
        Array.map (wrapErrorValues >> Option.bind validationfunction)

    /// Transforms a 1D array into an Array with typed (casted) elements
    /// Inputs: validationfunction validates the elements of the array elementswise after they've been checked for error/missing values
    ///         defaultvalue is the default value to replace missing values with
    ///         arr is the array to transform
    let array1DAsArrayWithDefault validationfunction defaultvalue =
        Array.map ( wrapErrorValues >>
                    Option.bind validationfunction >>
                    (FSharpx.Option.getOrElse defaultvalue))

    /// Transforms a 2D array into an Array2D with typed (casted) elements and the data elements in the various rows.
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenets of the array elementwise after they've been checked for error/missing values
    ///         arr is the array to transform
    let array2DAsArray roworiented (validationfunction : 'a -> 'b option) arr =
        arr |> fun x -> if roworiented then x else Array2D.transpose x
            |> Array2D.map (wrapErrorValues >> Option.bind validationfunction)

    /// Transforms a 2D array into an Array2D with typed (casted) elements and the data elements in the various rows. Errors are replaced by a defaultvalue
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         defaultvalue is the value elements get replaced with if the original value is an error
    ///         arr is the array to transform    
    let array2DAsArrayWithDefault roworiented validationfunction defaultvalue arr =
        arr |> array2DAsArray roworiented validationfunction
            |> Array2D.map (defaultArg defaultvalue)

    /// Transforms a 2D array into an sequence with typed (casted) elements and the data elements in the various rows.
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         arr is the array to transform  
    let array2DAsSeq roworiented validationfunction arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsSequence validationfunction

    /// Transforms a 2D array into an sequence with typed (casted) elements and the data elements in the various rows. Errors are replaced by a defaultvalue
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         defaultvalue is the value elements get replaced with if the original value is an error
    ///         arr is the array to transform  
    let array2DAsSeqWithDefault roworiented validationfunction defaultvalue transform arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsSequenceWithDefault validationfunction defaultvalue
                                                        
    /// Transforms a 2D array into an (1D) List with typed (casted) elements and the data elements in the various rows.
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         arr is the array to transform  
    let array2DAsList roworiented transform validationfunction  arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsList validationfunction            
 
    /// Transforms a 2D array into an (1D) List with typed (casted) elements and the data elements in the various rows. Errors are replaced by a defaultvalue
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         defaultvalue is the value elements get replaced with if the original value is an error
    ///         arr is the array to transform  
    let array2DAsListWithDefault roworiented transform validationfunction defaultvalue  arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsListWithDefault validationfunction defaultvalue   
            
    /// Transforms a 2D array into an 1D Array with typed (casted) elements and the data elements in the various rows.
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the lemenet so fhte array elementwise after they've been checked for error/missing values
    ///         arr is the array to transform 
    let array2DAsArray1D roworiented transform validationfunction arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsArray validationfunction
            
    /// Transforms a 2D array into an (1D) Array with typed (casted) elements and the data elements in the various rows. Errors are replaced by a defaultvalue
    /// Inputs: roworiented defines whether the array is interpreted as having rows contain different data elements. If false, the array gets transposed
    ///         validationfunction validates the elements of the array elementwise after they've been checked for error/missing values
    ///         defaultvalue is the value elements get replaced with if the original value is an error
    ///         arr is the array to transform  
    let array2DAsArray1DWithDefault roworiented transform validationfunction defaultvalue  arr =
        let arr = if roworiented then arr else Array2D.transpose arr
        [| for r in 0 .. (Array2D.length1 arr) - 1 do
                yield [| for c in 0 .. (Array2D.length2 arr) - 1 do yield arr.[r,c] |] |]
            |> array1DAsArrayWithDefault validationfunction defaultvalue   

    // array to return if an empty aray has to be returned
    let private empty2DArray = Array2D.init 1 1 (fun r c -> valueToExcel None)


    /// Transforms a sequence of generic data into an 2D array that can be returned to excel as UDF output
    /// asColumnVector is a boolean that specifies whether the elements of the data sequences are to be
    /// interpreted as row elements (if asColumnVector = true) or column elements (asColumnVector=false)
    /// data is a sequence of elements that need to be outputted.
    /// Transform is a transformation to be applied to the data. There are two things that this funtion
    /// must do. It must return an array; the various elements of this array are seen as the column (row)
    /// elements of the output array (which depends on asColumnVector). The second requirement is that
    /// every element in this array is an option type. Values that equal None are returned as error values
    /// to excel.
    let public array2DToExcel asColumnVector data transform =
        // Not a very efficient implementation as the sequence "data" gets traversed multiple times
        // and some superfluous array initialization takes place that probably can be avoided at the
        // cost of slightly more verbose code. However, until performance becomes unacceptable, I
        // prefer the clarity and simplicity of this code to faster, but more verbose code.
        let nmbRows = Seq.length data
        if nmbRows = 0 then empty2DArray
        else
            let nmbCols = data |> Seq.nth 0
                               |> transform
                               |> Array.length

            let res = Array2D.create nmbRows nmbCols (valueToExcel None)
            do data |> Seq.iteri (fun row elem ->  elem |> transform
                                                        |> Array.map (Option.bind wrapprimitive)
                                                        |> Array.iteri (fun col value ->  res.[row,col] <- valueToExcel value))
            
            res |> fun x -> if asColumnVector then x else Array2D.transpose x
    
    /// Processes a sequence into an 2D array that can be returned to excel. This 2D array contains only one column and (potentially) multiple rows
    let public arrayToExcelColumn data = array2DToExcel true data (fun x -> [|Some(x)|] )
    
    /// Processes a sequence into an 2D array that can be returned to excel. This 2D array contains only one row and (potentially) multiple columns
    let public arrayToExcelRow data = array2DToExcel false data (fun x -> [|Some(x)|] )

    /// Processes a sequence of values wrapped in options into an 2D array that can be returned to excel. This 2D array contains only one column and (potentially) multiple rows
    let public optionArrayToExcelColumn data = array2DToExcel true data (fun x -> [|x|])
    
    /// Processes a sequence of values wrapped in options into an 2D array that can be returned to excel. This 2D array contains only one row and (potentially) multiple columns
    let public optionArrayToExcelRow data = array2DToExcel false data (fun x -> [|x|])