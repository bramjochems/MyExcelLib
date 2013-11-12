namespace BJExcelLib.Finance

// Defines a period as consisting of a start end enddate
type public Period = { startDate:System.DateTime ; endDate:System.DateTime}
    
// Describes a calendar; consists of
//    - a set describing what days compromise the weekend
//    - a set describing the holidays for the particular calendar
type public Calendar = { weekendDays:System.DayOfWeek Set; holidays:System.DateTime Set }
    
/// Defines a tenor, which is an offset to a date. Consists of
///   - a years field specifiying the year to offset
///   - a months field specifying the number of months as offset
///   - a days field specifying the number of days offset
type public Tenor = { years:int; months:int; days:int }

/// The various types of roll rules that are implemented
type public RollRule =
    | Unadjusted
    | Following
    | Previous
    | ModifiedFollowing
    | ModifiedPrevious

/// The various daycount conventions implemented
type public DaycountConvention =
    | Actual_365qrt
    | Actual_365
    | Actual_360
    | Actual_Actual_ISDA
    | Thirty_360_E
    | Thirty_360_ISDA
    | Thirty_360_Eplus

/// Module containing date related functionality.
/// This module is partially based on http://www.tryfsharp.org/create/lesscode/DiscountCurve.fsx
module Date =

    // Parses a date from a string if possible
    let public date = fun s ->  match System.DateTime.TryParse(s) with
                                    | false, _      ->  None
                                    | true, value   ->  Some(value)

    /// Function produces a new date from a tenor and a startdate
    /// Two input parameters: tenor = tenor to offset with
    ///                       date  = date to offset from
    let public offset tenor (date:System.DateTime) = date.AddDays(float tenor.days).AddMonths(tenor.months).AddYears(tenor.years)

    /// Function that produces true if a date is a business day in the given calendar and false otherwise
    let public isBusinessDay  calendar (date:System.DateTime) = not (calendar.weekendDays.Contains date.DayOfWeek || calendar.holidays.Contains date)

    /// Produces the date immediately following the given date
    let private dayAfter (date:System.DateTime) = date.AddDays(1.0)
        
    /// Produces the date immediately preceding the given date
    let private dayBefore (date:System.DateTime) = date.AddDays(-1.0)

    /// All businessdays between two dates (inclusive) given a calendar
    let public businessDaysBetween calendar (startDate : System.DateTime) (endDate : System.DateTime) =
        let rec builder currDate enddate acc =
            if currDate > enddate then acc
            else
                if isBusinessDay calendar currDate then builder (dayAfter currDate) endDate (currDate :: acc)
                else builder (dayAfter currDate) endDate acc
        builder (min startDate endDate) (max startDate endDate) []
            |> List.rev

    /// Produces the nearest business day to a date given a specific roll rule to
    /// apply and given a calendar that defines business days. In particular, if
    /// a day already is a business day, it doesn't get rolled.
    /// Inputs are: a rule (type RollRule) that defines the rolling rule
    ///             a calendar that defines the business days
    ///             a date to start rolling from
    let rec public roll rule calendar date =
        if isBusinessDay calendar date then date
        else match rule with
                | RollRule.Unadjusted           ->  date
                | RollRule.Following            ->  date |> dayAfter |> roll rule calendar
                | RollRule.Previous             ->  date |> dayBefore|> roll rule calendar
                | RollRule.ModifiedFollowing    ->
                        let next = roll RollRule.Following calendar date
                        if next.Month <> date.Month then roll RollRule.Previous calendar date else next
                    | RollRule.ModifiedPrevious    ->
                        let prev = roll RollRule.Previous calendar date
                        if prev.Month <> date.Month then roll RollRule.Following calendar date else prev

    /// Rolls a specific date by n days, taking into a account a given roll rule and a
    /// calendar defining business days. If we're rolling 0 days or when we're using
    /// RollRule.Actual then it's possible to end up on a non-business days.
    /// Inputs are:     n - the number of days to roll
    ///                 rule - the rolling rule to apply
    ///                 calendar - the calendar that defines business days
    ///                 date     - the startdate from which we're rolling
    /// Note that rolling backwards in time is achieved by applying rollBy to a positive
    /// number of days to roll, but applying a (Modified) Previous roll rule.
    let rec public rollBy n rule calendar (date:System.DateTime) =
        match n with | 0 ->  date
                     | x ->  match rule with
                                | RollRule.Unadjusted           ->  date.AddDays(float x)
                                | RollRule.Following            ->  date|> dayAfter
                                                                        |> roll rule calendar
                                                                        |> rollBy (x - 1) rule calendar

                                | RollRule.Previous             ->  date|> roll rule calendar
                                                                        |> dayBefore
                                                                        |> roll rule calendar
                                                                        |> rollBy (x - 1) rule calendar

                                | RollRule.ModifiedFollowing    ->  // Roll n-1 days Following
                                                                    let next = rollBy (x - 1) RollRule.Following calendar date
                                                                    // Roll the last day ModifiedFollowing
                                                                    let final = roll RollRule.Following calendar (dayAfter next)
                                                                    if final.Month <> next.Month then (roll RollRule.Previous calendar next) else final
 
                                | RollRule.ModifiedPrevious     ->  // Roll n-1 days Previous
                                                                    let next = rollBy (x - 1) RollRule.Previous calendar date
                     
                                                                    // Roll the last day ModifiedPrevious
                                                                    let final = roll RollRule.Previous calendar (dayAfter next)
                                                                    if final.Month <> next.Month then roll RollRule.Following calendar next else final
    
    
    let private regex s = new System.Text.RegularExpressions.Regex(s,System.Text.RegularExpressions.RegexOptions.IgnoreCase)
      
    /// Produces a tenor from an input string. Input strings are concatenations
    /// of integers and upper case characters in the set "Y,M,W,D" denoting years
    /// months, weeks and days respectively. Not all combinations are parseable.
    /// To denote this, the result gets wrapped into an option type is it valid and
    /// invalid results are encoded by None.
    let public tenor (t : string)=
        let pattern = regex (    "(?<years>(-)?[0-9]+)Y(?<months>(-)?[0-9]+)M(?<weeks>(-)?[0-9]+)W(?<days>(-)?[0-9]+)D" +
                                    "|(?<years>(-)?[0-9]+)Y(?<months>(-)?[0-9]+)M(?<weeks>(-)?[0-9]+)W" +
                                    "|(?<years>(-)?[0-9]+)Y(?<months>(-)?[0-9]+)M(?<days>(-)?[0-9]+)D" +
                                    "|(?<years>(-)?[0-9]+)Y(?<months>(-)?[0-9]+)M" + 
                                    "|(?<years>(-)?[0-9]+)Y(?<weeks>(-)?[0-9]+)W(?<days>(-)?[0-9]+)D" +
                                    "|(?<years>(-)?[0-9]+)Y(?<weeks>(-)?[0-9]+)W" +
                                    "|(?<years>(-)?[0-9]+)Y(?<days>(-)?[0-9]+)D" +
                                    "|(?<years>(-)?[0-9]+)Y" +
                                    "|(?<months>(-)?[0-9]+)M(?<weeks>(-)?[0-9]+)W(?<days>(-)?[0-9]+)D" +
                                    "|(?<months>(-)?[0-9]+)M(?<weeks>(-)?[0-9]+)W" +
                                    "|(?<months>(-)?[0-9]+)M(?<days>(-)?[0-9]+)D" +
                                    "|(?<months>(-)?[0-9]+)M" +
                                    "|(?<weeks>(-)?[0-9]+)W(?<days>(-)?[0-9]+)D"+
                                    "|(?<weeks>(-)?[0-9]+)W" +
                                    "|(?<days>(-)?[0-9]+)D" )

        let m = pattern.Match(t.Replace(" ",""))
        if m.Success then
            Some( { years  = (if m.Groups.["years"].Success then int m.Groups.["years"].Value else 0);
                    months = (if m.Groups.["months"].Success then int m.Groups.["months"].Value else 0);
                    days   = match (m.Groups.["weeks"].Success), (m.Groups.["days"].Success) with
                                | true,  true   ->  7*(int m.Groups.["weeks"].Value ) + (int m.Groups.["days"].Value)
                                | true,  false  ->  7*(int m.Groups.["weeks"].Value ) 
                                | false, true   ->                                      (int m.Groups.["days"].Value)
                                | false, false  ->  0})        
            else None

    let private monthStringToNumber (s : string) =
        match s.ToLower() with
                | "jan" | "f" -> 1  |> Some
                | "feb" | "g" -> 2  |> Some
                | "mar" | "h" -> 3  |> Some
                | "apr" | "j" -> 4  |> Some
                | "may" | "k" -> 5  |> Some
                | "jun" | "m" -> 6  |> Some
                | "jul" | "n" -> 7  |> Some
                | "aug" | "q" -> 8  |> Some
                | "sep" | "u" -> 9  |> Some
                | "oct" | "v" -> 10 |> Some
                | "nov" | "x" -> 11 |> Some
                | "dec" | "z" -> 12 |> Some
                | _           ->       None

    /// Produces a period from an input string. Examples of inputs recognized are:
    /// Cal16, Q413, 13Q4, 3Q15, H114, 14H1, 1H14 Jan15, F14, H14-K14
    /// This function cannot recogize periods that are composed of multiple periods, concatenated by a dash.
    /// e.g. Cal16-Cal17 or 13Q1-Q4, H214-Cal15, Jan16-Dec18 all do NOT work
    let public period (s : string) =
        let s = s.Trim().Replace(" ","").Replace("-","").Replace("/","").Replace(@"\","")

        let fallthrough (y : 'a option) (x : 'a option) = // Probably should be refactored to some generic option functionality 
            if x.IsSome then x else y
            
        let yearpattern = regex "^cal(?<year>\d{2})$|^cal(?<year>\d{2})$"
        let semiyearpattern = regex "^H(?<period>[1-2])(?<year>\d{2})$|^(?<year>\d{2})H(?<period>[1-2])$|^(?<period>[1-2])H(?<year>\d{2})$|^H(?<period>[1-2])(?<year>\d{4})$|^(?<year>\d{4})H(?<period>[1-2])$|^(?<period>[1-2])H(?<year>\d{4})$"
        let quarterpattern = regex "^Q(?<period>[1-4])(?<year>\d{2})$|^(?<year>\d{2})Q(?<period>[1-4])$|^(?<period>[1-4])Q(?<year>\d{2})$|^Q(?<period>[1-4])(?<year>\d{4})$|^(?<year>\d{4})Q(?<period>[1-4])$|^(?<period>[1-4])Q(?<year>\d{4})$"
        let monthpattern1 = regex "^(?<period>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year>\d{2})$|^(?<period>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year>\d{4})$"
        let monthpattern2 = regex "^(?<period1>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year1>\d{2})(?<period2>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year2>\d{2})$|^(?<period1>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|okt|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year1>\d{2})(?<period2>(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|okt|nov|dec|f|g|h|j|k|m|n|q|u|v|x|z))(?<year2>\d{2})$"

        // todo : all logic regarding baseyear won't be working anymore once we get near to the turn of a century
        let baseyear = System.DateTime.Today.Year - (System.DateTime.Today.Year % 100)

        let y = yearpattern.Match(s)
        if y.Success then
            let year = int y.Groups.["year"].Value
            let year = if year < 100 then year + baseyear else year
            Some({startDate = System.DateTime(year,1,1);endDate = System.DateTime(year,12,31)})
        else
            None
            |> fallthrough (    let h = semiyearpattern.Match(s)
                                if h.Success then
                                    let half = int h.Groups.["period"].Value
                                    let year = int h.Groups.["year"].Value
                                    let year = if year < 100 then year + baseyear else year
                                    let start = System.DateTime(year,1+(half-1)*6,1)
                                    Some({startDate = start; endDate = start.AddMonths(6).AddDays(-1.)})
                                else None)
            |> fallthrough (    let q = quarterpattern.Match(s)
                                if q.Success then
                                    let half = int q.Groups.["period"].Value
                                    let year = int q.Groups.["year"].Value
                                    let year = if year < 100 then year + baseyear else year
                                    let start = System.DateTime(year,1+(half-1)*3,1)
                                    Some({startDate = start; endDate = start.AddMonths(3).AddDays(-1.)})
                                else None)
            |> fallthrough (    let m = monthpattern1.Match(s)
                                if m.Success then
                                    let month = m.Groups.["period"].Value |> monthStringToNumber
                                    if month.IsSome then
                                        let year = int m.Groups.["year"].Value
                                        let year = if year < 100 then year + baseyear else year
                                        let start = System.DateTime(year,month.Value,1)
                                        Some({startDate = start; endDate = start.AddMonths(1).AddDays(-1.)})
                                    else None
                                else None)
                                     
            |> fallthrough (    let m = monthpattern2.Match(s)
                                if m.Success then
                                    let month1 = m.Groups.["period1"].Value |> monthStringToNumber
                                    let month2 = m.Groups.["period2"].Value |> monthStringToNumber
                                    if (month1.IsSome) && (month2.IsSome) then
                                        let year1 = int m.Groups.["year1"].Value
                                        let year2 = int m.Groups.["year2"].Value
                                        let year1 = if year1 < 100 then year1 + baseyear else year1
                                        let year2 = if year2 < 100 then year2 + baseyear  else year2
                                        let start = System.DateTime(year1,month1.Value,1)
                                        let endd = System.DateTime(year2,month2.Value,1).AddMonths(1).AddDays(-1.)
                                        if start < endd then Some({startDate =  start; endDate=endd}) else None
                                    else None
                                else None)

    /// Calculations the fraction of time between startdate and enddate using a specified daycount convention
    let public yearfrac daycount (startdate : System.DateTime ) (enddate : System.DateTime)=
        
        // Helper method for calcultions with 30/360 methods
        let helper30360 (y1,m1,d1) (y2,m2,d2) =
            let y1,y2,m1,m2,d1,d2 = float y1, float y2, float m1, float m2, float d1, float d2
            (360.*(y2-y1) + 30.*(m2-m1) + d2-d1)/360.
        let islastdayofmonth (d : System.DateTime) = d.AddDays(1.).Month <> d.Month
        let isleapyear (d : System.DateTime) = System.DateTime.IsLeapYear(d.Year)

        // Main logic
        let helper dc (startd : System.DateTime)  (endd : System.DateTime) =
            match dc with
                | Actual_365qrt         ->  (endd.Subtract(startd).TotalDays + 1.)/365.25
                | Actual_365            ->  (endd.Subtract(startd).TotalDays + 1. )/365.
                | Actual_360            ->  (endd.Subtract(startd).TotalDays + 1. )/360.
                | Actual_Actual_ISDA    ->  match endd.Year = startd.Year with
                                                | true  ->  (endd.Subtract(startd).TotalDays + 1.)/(if isleapyear startd then 366. else 365.)
                                                | false ->  let addon = max ((float) (endd.Year-startd.Year) - 1.) 0.
                                                            let frac1 = (System.DateTime(startd.Year,12,31).Subtract(startd).TotalDays + 1.) /(if isleapyear startd then 366. else 365.)
                                                            let frac2 = (endd.Subtract(System.DateTime(endd.Year,1,1)).TotalDays + 1.)  / (if isleapyear endd then 366. else 365.)
                                                            addon+frac1+frac2        
                | Thirty_360_E          ->  helper30360 (startd.Year, startd.Month, min startd.Day 30) (endd.Year, endd.Month, min endd.Day 30)
                | Thirty_360_ISDA       ->  let startdata = (startd.Year, startd.Month, if (islastdayofmonth startd) then 30 else startd.Day)
                                            let enddata = (endd.Year, endd.Month, if endd.Month <> 2 && (islastdayofmonth endd) then 30 else endd.Day)
                                            helper30360 startdata enddata
                | Thirty_360_Eplus      ->  if endd.Day = 31 then helper30360 (startd.Year,startd.Month,min startd.Day 30) (endd.Year,endd.Month+1,1)
                                            else helper30360 (startd.Year,startd.Month,min startd.Day 30) (endd.Year,endd.Month,endd.Month)
        
        // Deal with input order
        if startdate = enddate then 0.
        elif startdate > enddate then -(helper daycount enddate startdate)
        else helper daycount startdate enddate