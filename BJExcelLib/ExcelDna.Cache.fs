namespace BJExcelLib.ExcelDna

open ExcelDna.Integration
open ExcelDna.Integration.Rtd
open System.Collections.Generic


/// Module that allows creation of persistent objects from excel. Got the idea 
/// (and a big part of the code) from the ExcelDNA group in a post by "DC":
/// https://groups.google.com/forum/#!searchin/exceldna/xlcache/exceldna/E4vIOSNwHm0/yHmVTW4FHA0J
module public Cache =

    /// The cache that holds objects indexed by a key
    let private cache = Dictionary<_,_>()
    let private tagstore = Dictionary<_,_>()

    /// Registers an object in the cache. The tag is an id for the category the object belongs to
    /// The tag is not equal to the final handle the object gets, because a counter will be prepared
    let public register tag o =
        let counter = 
            match tagstore.TryGetValue tag with
                | true, c -> c + 1
                | _       -> 1

        tagstore.[tag] <- counter
        let handle = tag + "." + counter.ToString()
        cache.[handle] <- box o
        XlCall.RTD("BJExcelLib.ExcelDna.CacheRTD", null, handle)

    /// Removes a given handle from the cache
    let public unregister handle =    
        if cache.ContainsKey(handle) then cache.Remove(handle) |> ignore

    /// resets the cache by removing all values in there
    let public reset =
        cache.Clear()
        tagstore.Clear()

    /// Finds a value in the cache. If no value is found, None is returned. If a value is found, then
    /// the the result is wrapped in Some and returned, otherwise None is returned.
    let lookup handle =
        match cache.TryGetValue handle with
            | true, value -> Some(value)
            | _           -> None

/// Excel RTD server to handle the registering/unregistering of object handles
type public CacheRTD() =
    inherit ExcelRtdServer() 
    let _topics = new Dictionary<ExcelRtdServer.Topic, string>()
    override x.ConnectData(topic:ExcelRtdServer.Topic, topicInfo:IList<string>, newValues:bool byref) =
            let name = topicInfo.Item(0)
            _topics.[topic] <- name
            name |> box
    override x.DisconnectData(topic:ExcelRtdServer.Topic) =
                _topics.[topic] |> Cache.unregister
                _topics.Remove(topic) |> ignore