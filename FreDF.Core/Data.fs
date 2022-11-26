namespace FreDF.Core

[<RequireQualifiedAccess>]
module Data =

    open System.Text.Json
    open System.Text.RegularExpressions

    module Internal =

        /// Captures a path part like in the format `key[iterator]`.
        let pathPartRegex =
            Regex(
                "^^(?<key>.*?)\[(?<iterator>.*?)\]$$",
                RegexOptions.ExplicitCapture
                ||| RegexOptions.Compiled
            )

    type Context =
        { Data: Data
          Iterators: Iterators }

        member ctx.AdvanceIterator(name: string) =
            { ctx with Iterators = ctx.Iterators.Advance name }

        member ctx.ResolveValue(path: Path) =
            path.Parts
            |> List.fold
                (fun curr part ->
                    curr
                    |> Option.bind (fun value ->
                        match part with
                        | PathPart.Array (key, iteratorName) ->
                            // If array then fetch array from object
                            // and get from current iterator index.
                            match value with
                            | Value.Object o ->
                                o.TryFind key
                                |> Option.bind (function
                                    | Value.Array a -> Some a
                                    | _ -> None)
                                |> Option.bind (fun a ->
                                    ctx.Iterators.GetIndex iteratorName
                                    |> Option.bind (fun i -> a |> List.tryItem i))
                            | _ -> None
                        | PathPart.Object key ->
                            // If object then fetch from object.
                            match value with
                            | Value.Object o -> o.TryFind key
                            | _ -> None))
                (ctx.Data.AsValue() |> Some)

        member ctx.ResolveValueToScalar(path: Path) =
            ctx.ResolveValue path
            |> Option.bind Value.TryToScalar

        member ctx.ResolveIterator(path: Path) =
            // Go from root.
            // traverse path to find last array
            // Return length.
            let r =
                path.Parts
                |> List.fold
                    (fun (curr: Value option) p ->
                        curr
                        |> Option.bind (fun v ->
                            match p with
                            | PathPart.Array (key, _) ->
                                match v with
                                | Value.Object o ->
                                    o.TryFind key
                                    |> Option.bind (function
                                        | Value.Array a -> Some(Value.Array a)
                                        | _ -> None)
                                | Value.Array a ->
                                    // If the value is an array...
                                    // fetch the first item as an object and then try and find the iterator key in that.
                                    // This does mean if the second or third items (etc.) are different,
                                    // this will not reflected.
                                    // However there is probably no simple way to handle objects which could be different.
                                    // This first item possible could be checked to make sure it is an array.
                                    a
                                    |> List.tryItem 0
                                    |> Option.bind (function
                                        | Value.Object o -> o.TryFind key
                                        | _ -> None)
                                | _ -> None
                            | PathPart.Object key ->
                                match v with
                                | Value.Object o -> o.TryFind key
                                | _ -> None))
                    (ctx.Data.AsValue() |> Some)

            r
            |> Option.bind (function
                | Value.Array v -> Some v.Length
                | _ -> None)

        member ctx.AddIterator(name: string, path: Path) =
            { ctx with Iterators = ctx.Iterators.Add(name, path) }

        member ctx.RemoveIterator(name: string) =
            { ctx with Iterators = ctx.Iterators.Remove(name) }
    (*
            let root =
                match path.Root with
                | PathRoot.Root ->
                    // If root, get the root data as an object.
                    ctx.Data.AsValue() |> Some
                | PathRoot.Iterator name ->
                    // If a named iterator,
                    // traverse back to root to get the current value.

                    //let traverse =


                    None
            //| PathPart.O
            *)


    and Data =
        { Values: Map<string, Value> }

        static member FromJson(element: JsonElement) =
            let rec build (el: JsonElement) =
                match el.ValueKind with
                | JsonValueKind.Array ->
                    el.EnumerateArray()
                    |> List.ofSeq
                    |> List.choose build
                    |> Value.Array
                    |> Some
                | JsonValueKind.Object ->
                    el.EnumerateObject()
                    |> List.ofSeq
                    |> List.choose (fun p -> build p.Value |> Option.map (fun v -> p.Name, v))
                    |> Map.ofList
                    |> Value.Object
                    |> Some
                | JsonValueKind.String -> el.GetString() |> Value.Scalar |> Some
                | JsonValueKind.Number -> None
                | JsonValueKind.True -> Value.Scalar "True" |> Some
                | JsonValueKind.False -> Value.Scalar "False" |> Some
                | JsonValueKind.Null -> Value.Scalar "" |> Some
                | JsonValueKind.Undefined -> None
                | _ -> None

            match element.ValueKind with
            | JsonValueKind.Object ->
                element.EnumerateObject()
                |> List.ofSeq
                |> List.choose (fun p -> build p.Value |> Option.map (fun v -> p.Name, v))
                |> Map.ofList
                |> fun vs -> { Values = vs }
                |> Ok
            | _ -> Error $"Unsupported top level value: `{element.ValueKind}`. Only objects are currently supported."

        member d.AsValue() = d.Values |> Value.Object

    and [<RequireQualifiedAccess>] Value =
        | Scalar of string
        | Object of Map<string, Value>
        | Array of Value list

        static member TryToScalar(v: Value) =
            match v with
            | Scalar s -> Some s
            | _ -> None

        member v.GetScalar() =
            match v with
            | Scalar s -> Some s
            | _ -> None

    and PathPart =
        | Object of Key: string
        | Array of Key: string * IteratorName: string

    and PathRoot =
        | Root
        | Iterator of Name: string

    and Path =
        { Root: PathRoot
          Parts: PathPart list }

        static member Parse(value: string) =
            let parts = value.Split('.')

            let root =
                match parts |> Array.tryHead with
                | Some "$" -> PathRoot.Root
                | Some root -> PathRoot.Iterator root
                | None -> PathRoot.Root

            let parsePath (value: string) =
                let matches = Internal.pathPartRegex.Match value
                
                match matches.Success with
                | true ->
                    // Check this doesn't fail!
                    PathPart.Array(matches.Groups["key"].Value, matches.Groups["iterator"].Value)
                | false -> PathPart.Object value


            { Root = root
              Parts =
                parts
                |> Array.tail
                |> Array.map parsePath
                |> List.ofArray }



    and Iterators =
        { Items: Map<string, Iterator> }

        member is.Add(name: string, path: Path) =
            { is with Items = is.Items.Add(name, Iterator.Create path) }

        member is.Remove(name: string) =
            let rec traverse (name: string, state: Map<string, Iterator>) =
                let newState = state.Remove name

                let iterators =
                    newState
                    |> Map.filter (fun k v -> v.Parent = Some name)

                match iterators.IsEmpty with
                | true -> newState
                | false ->
                    iterators
                    |> Map.fold (fun s k v -> traverse (k, s)) newState

            { is with Items = traverse (name, is.Items) }

        member is.Advance(name: string) =
            is.Items.TryFind name
            |> Option.map (fun i -> i, is.Items.Add(name, i.Next()))

            //{ is with Items = is.Items.Add(name, i.Next()) })
            |> Option.map (fun (iterator, newState) ->
                let rec traverse (name: string, state: Map<string, Iterator>) =
                    // This should look for any children (not parents)...

                    state
                    |> Map.filter (fun _ v -> v.Parent = Some name)
                    |> Map.fold
                        (fun (state: Map<string, Iterator>) k v ->
                            state.Add(k, v.Reset())
                            |> fun s -> traverse (k, s))
                        (state)

                (*
                    match state.TryFind name with
                    | Some i ->
                        match i.Parent with
                        | Some parent -> traverse (parent, state.Add(name, i.Reset()))
                        | None -> state.Add(name, i.Reset())
                    | None -> state
                    *)
                traverse (name, newState)

            (*
                match iterator.Parent with
                | Some parent -> traverse (parent, newState)
                | None -> newState*) )
            |> Option.map (fun newState -> { is with Items = newState })
            |> Option.defaultValue is

        member is.GetIndex(name: string) =
            is.Items.TryFind name
            |> Option.map (fun i -> i.CurrentIndex)

        member is.GetPathPaths(name: string) =

            // Back track up the iterator chain until reaching one with root as path type
            let rec traverse (iteratorName: string, parts: PathPart list) =
                match is.Items.TryFind iteratorName with
                | Some i ->
                    match i.Path.Root with
                    | PathRoot.Root -> i.Path.Parts @ parts
                    | PathRoot.Iterator n -> traverse (n, i.Path.Parts @ parts)
                | None -> parts

            traverse (name, [])

        member is.GetFullPath(name: string) =
            // Back track up the iterator chain until reaching one with root as path type

            let rec traverse (iteratorName: string, parts: PathPart list) =
                match is.Items.TryFind iteratorName with
                | Some i ->
                    match i.Path.Root with
                    | PathRoot.Root -> i.Path.Parts @ parts
                    | PathRoot.Iterator n -> traverse (n, i.Path.Parts @ parts)
                | None -> parts

            { Root = PathRoot.Root
              Parts = traverse (name, []) }

        member is.ExpandPath(path: Path) =
            match path.Root with
            | PathRoot.Root -> path
            | PathRoot.Iterator name ->
                // Any extra paths

                { Root = PathRoot.Root
                  Parts = is.GetPathPaths name @ path.Parts }

    and Iterator =
        { Parent: string option
          Path: Path
          CurrentIndex: int }

        static member Create(path: Path) =
            { Parent =
                path.Root
                |> function
                    | PathRoot.Iterator n -> Some n
                    | _ -> None
              Path = path
              CurrentIndex = 0 }

        member i.Reset() = { i with CurrentIndex = 0 }

        member i.Next() =
            { i with CurrentIndex = i.CurrentIndex + 1 }
