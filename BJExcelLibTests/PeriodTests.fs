namespace BJExcelLibTests

open System
open Microsoft.VisualStudio.TestTools.UnitTesting
open BJExcelLib.Finance
open BJExcelLib.Finance.Date


[<TestClass>]
type PeriodTests() =

    let testdata =
        [ ("Cal15", System.DateTime(2015,1,1), System.DateTime(2015,12,31));
          ("Q413", System.DateTime(2013,10,1), System.DateTime(2013,12,31));
          ("2023H1", System.DateTime(2023,1,1), System.DateTime(2023,6,30));
          ("F15-Z15", System.DateTime(2015,1,1), System.DateTime(2015,12,31));
          ("14Q2", System.DateTime(2014,4,1), System.DateTime(2014,6,30));
          ("Aug16", System.DateTime(2016,8,1), System.DateTime(2016,8,31));
          ("Q18", System.DateTime(2018,8,1), System.DateTime(2018,8,31));
          ("Q18-Q19", System.DateTime(2018,8,1), System.DateTime(2019,8,31));
          ("H14G15", System.DateTime(2014,3,1), System.DateTime(2015,2,28));
          ("H12", System.DateTime(2012,3,1), System.DateTime(2012,3,31)) ]

    [<TestMethod>]
    member x.``Period function returns the correct periods``() =
        let test = testdata |> List.map(fun (t, s, e) -> let per = period t
                                                         per.IsSome && per.Value.startDate = s && per.Value.endDate = e)
                            |> List.fold (fun acc v -> acc && v) true
        Assert.AreEqual(true,test)

