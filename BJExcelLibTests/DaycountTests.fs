namespace BJExcelLibTests

open System
open Microsoft.VisualStudio.TestTools.UnitTesting
open BJExcelLib.Finance
open BJExcelLib.Finance.Date


[<TestClass>]
type DaycountTests() =

    let testperiod =
        [   { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) };
            { startDate = System.DateTime(2010,1,1); endDate = System.DateTime(2010,12,31) }    ]

    [<TestMethod>]
    member x.``Actual/Actual (ISDA) convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.] //todo
        let test = testperiod   |> List.map (fun x -> yearfrac Actual_Actual_ISDA x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)


    [<TestMethod>]
    member x.``30/360 (ISDA) convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.]//todo
        let test = testperiod   |> List.map (fun x -> yearfrac Thirty_360_ISDA x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``30E/360 convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.]//todo
        let test = testperiod   |> List.map (fun x -> yearfrac Thirty_360_E x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``30E+/360 convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.]//todo
        let test = testperiod   |> List.map (fun x -> yearfrac Thirty_360_Eplus x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Act/360 convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.]//todo
        let test = testperiod   |> List.map (fun x -> yearfrac Actual_360 x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)

    [<TestMethod>]
    member x.``Act/365 convention results are correct`` () =
        let expected = [0.;
                        1.;
                        2.;
                        3.;
                        4.;
                        5.;
                        6.;
                        7.;
                        8.;
                        9.]//todo
        let test = testperiod   |> List.map (fun x -> yearfrac Actual_365 x.startDate x.endDate)
                                |> List.zip expected
                                |> List.fold (fun acc (act,exp) -> acc && (act = exp)) true
        Assert.AreEqual(true,test)

