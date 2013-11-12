namespace BJExcelLib.UDF

/// Contains excel UDFs for date related functions
module public Date =
    open ExcelDna.Integration
    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.Finance
    open BJExcelLib.Finance.Date

    /// Non-volatile version of Excel's today() function
    [<ExcelFunction(Name="Today_Nonvolatile", Description="Today's date as a non-volatile function",Category="BJFinanceLib.Date", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public today_nonvolatile = fun () -> System.DateTime.Today

    /// Number of calendar days in the calendar month specified by a given date
    [<ExcelFunction(Name="CalendarDaysInMonth", Description="Number of calendar days in the month specified by a given date",Category="BJExcelLib.Date", IsThreadSafe = true,IsExceptionSafe=true)>]
    let public calendardaysinmonth ([<ExcelArgument(Name="Date",Description = "The date to get the number of calendar days for")>] date : obj) =
        date    |> validateDate
                |> Option.bind(fun x -> let thisMonthStart = new System.DateTime(x.Year, x.Month,1)
                                        thisMonthStart.AddMonths(1).Subtract(thisMonthStart).Days |> Some)
                |> valueToExcel

    /// Computes an enddate from a tenor string
    [<ExcelFunction(Name="DateFromTenor",Description="Calculates an enddate given a tenor and a startdate",Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public datefromtenor    ([<ExcelArgument(Name="Tenor",Description = "The tenor to apply to the startdate, e.g. '4y3d'")>] tenor : obj,
                                 [<ExcelArgument(Name="StartDate",Description = "Startdate. Optional, default = today")>] startdate : obj) =
    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse (System.DateTime.Today)
        tenor   |> validateString 
                |> Option.bind BJExcelLib.Finance.Date.tenor
                |> Option.bind (fun t -> (offset t startdate) |> Some)
                |> valueToExcel

    /// Computes a year fraction between two dates on an Act/Act (ISDA) daycount
    [<ExcelFunction(Name="TTM",Description="Calculations the yearfraction between two dates assuming an Act/Act ISDA convention", Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public TTM( [<ExcelArgument(Name="Enddate",Description = "Enddate of the ttm calculation (inclusive)")>] enddate : obj,
                    [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse System.DateTime.Today
        enddate |> validateDate
                |> Option.map (fun e -> yearfrac Actual_Actual_ISDA startdate e)
                |> valueToExcel

    /// Computes a year fraction between two dates on an 30/360 (ISDA) daycount
    [<ExcelFunction(Name="TTM_30_360",Description="Calculations the yearfraction between two dates assuming an 30/360 ISDA convention", Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public TTM30360([<ExcelArgument(Name="Enddate",Description = "Enddate of the ttm calculation (inclusive)")>] enddate : obj,
                        [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse System.DateTime.Today
        enddate |> validateDate
                |> Option.map (fun e -> yearfrac Thirty_360_ISDA startdate e)
                |> valueToExcel

    /// Computes a year fraction between two dates on an Act/360 daycount
    [<ExcelFunction(Name="TTM_Act_360",Description="Calculations the yearfraction between two dates assuming an Act/360 convention", Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public TTMAct360([<ExcelArgument(Name="Enddate",Description = "Enddate of the ttm calculation (inclusive)")>] enddate : obj,
                         [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse System.DateTime.Today
        enddate |> validateDate
                |> Option.map (fun e -> yearfrac Actual_360 startdate e)
                |> valueToExcel

    /// Computes a year fraction between two dates on an  Act/365 daycount
    [<ExcelFunction(Name="TTM_Act_365",Description="Calculations the yearfraction between two dates assuming an Act/365.25 convention", Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public TTMAct365([<ExcelArgument(Name="Enddate",Description = "Enddate of the ttm calculation (inclusive)")>] enddate : obj,
                         [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse System.DateTime.Today
        enddate |> validateDate
                |> Option.map (fun e -> yearfrac Actual_365 startdate e)
                |> valueToExcel

        /// Computes a year fraction between two dates on an  Act/365 daycount
    [<ExcelFunction(Name="TTM_Act_365.25",Description="Calculations the yearfraction between two dates assuming an Act/365.25 convention", Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public TTMAct365qrt([<ExcelArgument(Name="Enddate",Description = "Enddate of the ttm calculation (inclusive)")>] enddate : obj,
                            [<ExcelArgument(Name="StartDate",Description = "Startdate of the ttm calculation (inclusive). Optional, default = today")>] startdate : obj) =    
        let startdate = startdate |> validateDate |> FSharpx.Option.getOrElse System.DateTime.Today
        enddate |> validateDate
                |> Option.map (fun e -> yearfrac Actual_365qrt startdate e)
                |> valueToExcel

    [<ExcelFunction(Name="Period.Startdate",Description="Calculates the startdate of a period denoted by a string",Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public PeriodStartDate([<ExcelArgument(Name="Period",Description="The period to get the startdate for, e.g. Cal16, Q413, 13H2, May15, etc")>] period : obj) =
        period  |> validateString
                |> Option.bind BJExcelLib.Finance.Date.period
                |> Option.map (fun z -> z.startDate)
                |> valueToExcel

    [<ExcelFunction(Name="Period.Enddate",Description="Calculates the startdate of a period denoted by a string",Category="BJExcelLib.Date",IsThreadSafe=true,IsExceptionSafe=true)>]
    let public PeriodEndDate([<ExcelArgument(Name="Period",Description="The period to get the startdate for, e.g. Cal16, Q413, 13H2, May15, etc")>] period : obj) =
        period  |> validateString
                |> Option.bind BJExcelLib.Finance.Date.period
                |> Option.map (fun z -> z.endDate)
                |> valueToExcel