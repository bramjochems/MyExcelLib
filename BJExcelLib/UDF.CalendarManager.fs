namespace BJExcelLib.UDF

/// This module contains UDFs to add and remove calendars via UDFs to a calendar manager (just a simple dictionary). Unlike with the ExcelCache
/// I am not using an RTD here since I don't want to deal with keeping track of which version of the object is available in excel. Also, lifetime
/// management is not much of an issue since I typically expect these kind of functionality to remain in memory during an entire excel session. Finally,
/// the most usage out of these UDFs is probably obtained by calling them in VBA code rather than in excel directly
module calendarManager =

    open BJExcelLib.ExcelDna.IO
    open BJExcelLib.Finance
    open BJExcelLib.Finance.Date
    open ExcelDna.Integration
    open System.Collections.Generic

    let calendarDict = new Dictionary<string,Calendar>()

    //Todo: automatic loading of calendar dictionary from file

    [<ExcelFunction(Name="CalendarManager.AddCalendar", Description="Adds a calendar to the calendar manager by specifying the holidays (weekend = Saturdays + Sundays)",Category="BJExcelLib.Date",IsExceptionSafe=true)>]
    let public addCalendar ([<ExcelArgument(Name="Name",Description = "Name of the calendar to add")>] name : obj,
                            [<ExcelArgument(Name="Jolidays",Description = "Holidays under the calendar")>] dates : obj []) =
        let name = name |> validateString |> Option.map (fun x -> x.ToUpper())
        if name.IsSome then
            let dates = dates |> array1DAsList validateDate
                              |> List.fold (fun acc x -> if x.IsSome then x.Value :: acc else acc) [] // filter and map in one
                              |> Set.ofList

            let we = [System.DayOfWeek.Saturday;System.DayOfWeek.Sunday] |> Set.ofList
            let cal = { weekendDays = we; holidays = dates }
            if calendarDict.ContainsKey(name.Value) then calendarDict.[name.Value] <- cal else calendarDict.Add(name.Value,cal)

        name |> valueToExcel

    [<ExcelFunction(Name="CalendarManager.RemoveCalendar", Description="Removes a calendar with a given name from the calendar database",Category="BJExcelLib.Date")>]
    let public removeCalendar ([<ExcelArgument(Name="Name",Description = "Name of the calendar to remove")>] name : obj) =
        name |> validateString
             |> Option.map (fun x -> x.ToUpper())
             |> Option.map (fun x -> calendarDict.Remove(x))
             |> valueToExcel

    [<ExcelFunction(Name="CalendarManager.GetCalendarNames", Description="Retrieves all available calendar names",Category="BJExcelLib.Date")>]
    let public getCalendarNames() =
        calendarDict.Keys   |> Seq.toArray
                            |> arrayToExcelColumn
                            
    [<ExcelFunction(Name="CalendarManager.GetHolidays", Description="Gets all holidays greater than startdate and smaller than enddate from the calendar",Category="BJExcelLib.Date")>]
    let public getCalendarHolidays ([<ExcelArgument(Name="Name",Description = "Name of the calendar for which to retrieve calendars")>] name : obj,
                                    [<ExcelArgument(Name="Startdate",Description = "Optional lower bound on holidays to retrieve")>] startDate : obj,
                                    [<ExcelArgument(Name="Enddate",Description = "Optional upper bound on holidays to retrieve")>] endDate : obj) =
        name |> validateString |> Option.map (fun x -> x.ToUpper())
             |> Option.map(fun calname ->   let calendar = if calendarDict.ContainsKey(calname) then calendarDict.[calname] else { weekendDays =Set.empty; holidays = Set.empty }
                                            let startDate = startDate |> validateDate
                                            let endDate = endDate |> validateDate
                                            calendar.holidays   |> Set.toList
                                                                |> fun x -> if startDate.IsSome then List.filter (fun date -> date >= startDate.Value) x else x
                                                                |> fun x -> if endDate.IsSome then List.filter (fun date -> date <= endDate.Value) x else x
                            )
             |> fun z -> if z.IsSome then z.Value else []
                            |> arrayToExcelColumn

    [<ExcelFunction(Name="CalendarManager.BusinessDaysBetween", Description="Gets all business days between a startdate and enddate for a given calendar",Category="BJExcelLib.Date")>]
    let public getCalendarBDays([<ExcelArgument(Name="Name",Description = "Name of the calendar for which to retrieve business days")>] name : obj,
                                [<ExcelArgument(Name="Startdate",Description = "start of the period")>] startDate : obj,
                                [<ExcelArgument(Name="Enddate",Description = "end of the period")>] endDate : obj) =
        
        let startDate = startDate |> validateDate
        let endDate = endDate |> validateDate
        if startDate.IsSome && endDate.IsSome then
            name    |> validateString |> Option.map (fun x -> x.ToUpper())
                    |> Option.map(fun calname -> if calendarDict.ContainsKey(calname) then calendarDict.[calname] else { weekendDays =Set.empty; holidays = Set.empty }
                                                    |> fun cal -> businessDaysBetween cal startDate.Value endDate.Value)
                    |> fun x -> x.Value
        else []
        |> List.toArray
        |> arrayToExcelColumn
                                