namespace FreDF.Core

open System.Text.Json
open System.Text.RegularExpressions
open MigraDocCore.DocumentObjectModel


module Templating =
    

    open ToolBox.Core

    
    module Internal =
        
        let contentRegex = Regex("\{\{(?<value>((\$\.|[A-Za-z]).+?))\}\}|\{\{%(?<preprocessor>(.+?))\}\}|(?<=}}|^)(?<literal>(.+?))(?={{|$)", RegexOptions.ExplicitCapture ||| RegexOptions.Compiled)
    
    [<RequireQualifiedAccess>]
    type Content =
        | Literal of string
        | Value of Data.Path
        | Processor

    type TemplatingContext =
        { Data: Data.Context
          Scopes: Scope list }

        member tc.ResolveValue(path: Data.Path) = tc.Data.ResolveValueToScalar(path)

        member tc.OpenScope(scope: Scope) =
            { tc with
                Data =
                    tc.Data.AddIterator(
                        scope.Name,
                        scope.Path
                    (*
                        tc.Scopes
                        |> List.tryHead
                        |> Option.map (fun s -> s.Name)
                        *)
                    )
                Scopes = scope :: tc.Scopes }

        member tc.CloseScope() =
            match tc.Scopes |> List.tryHead with
            | Some scope ->
                { tc with
                    Data = tc.Data.RemoveIterator(scope.Name)
                    Scopes = tc.Scopes.Tail }
            | None -> tc

        member tc.AdvanceIterator(name: string) =
            { tc with Data = tc.Data.AdvanceIterator name }

        member tc.Iterate(scope: Scope, fn: TemplatingContext -> 'T) =

            let ctx = tc.OpenScope(scope)

            ctx.Data.Iterators.GetFullPath scope.Name
            |> ctx.Data.ResolveIterator
            |> Option.map (fun i ->
                let rec iter (ctx: TemplatingContext, acc: 'T list) =
                    // Check if iterator has items left.
                    match ctx.Data.Iterators.GetIndex scope.Name with
                    | Some curr when curr < i -> iter (ctx.AdvanceIterator(scope.Name), fn ctx :: acc)
                    | Some _
                    | None -> ctx.CloseScope(), acc |> List.rev

                iter (ctx, []))
            |> Option.defaultValue (tc.CloseScope(), [])

    /// Scopes will attempt to create an iterator with the name.
    and Scope =
        { Name: string
          Path: Data.Path }

        static member TryParse(value: string) =
            let path = Data.Path.Parse(value)

            path.Parts
            |> List.tryLast
            |> Option.bind (function
                | Data.PathPart.Array (_, iterator) -> Some iterator
                | _ -> None)
            |> Option.map (fun name -> { Name = name; Path = path })

        static member TryFromJson(element: JsonElement) =
            match Json.tryGetStringProperty "name" element, Json.tryGetStringProperty "path" element with
            | Some name, Some path ->
                Some
                    { Name = name
                      Path =
                        { Root = Data.PathRoot.Root
                          Parts = [] } }
            | None, _
            | _, None -> None

    and SectionScopeType =
        | Standard
        | Inner

    and TemplatingResult<'T> =
        | Single of Result: 'T * Context: TemplatingContext
        | Multiple of Result: 'T list * Context: TemplatingContext

    type ConditionType = Bespoke of bool

    type Template =
        { Sections: TemplateSection list }

        static member FromJson(element: JsonElement) =
            { Sections =
                Json.tryGetArrayProperty "sections" element
                |> Option.map (List.map TemplateSection.FromJson)
                |> Option.defaultValue [] }

        member t.Build(ctx: TemplatingContext) =
            let sections =
                t.Sections
                |> List.fold
                    (fun (sections, ctx) s ->
                        match s.Build(ctx) with
                        | TemplatingResult.Single (result, newCtx) -> sections @ [ result ], newCtx
                        | TemplatingResult.Multiple (results, newCtx) -> sections @ results, newCtx)
                    ([], ctx)
                |> fst

            ({ Sections = sections }: Structure.PdfDocument)

    and TemplateSection =
        { Scope: Scope option
          ScopeType: SectionScopeType option
          PageSetup: Structure.PageSetup option
          Headers: HeaderFooterTemplates option
          Footers: HeaderFooterTemplates option
          Elements: DocumentElementTemplate list }

        static member FromJson(element: JsonElement) =
            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              ScopeType =
                Json.tryGetBoolProperty "innerScope" element
                |> Option.map (function
                    | true -> SectionScopeType.Inner
                    | false -> SectionScopeType.Standard)
              PageSetup =
                Json.tryGetProperty "pageSetup" element
                |> Option.map Structure.PageSetup.FromJson
              Headers =
                Json.tryGetProperty "headers" element
                |> Option.map HeaderFooterTemplates.FromJson
              Footers =
                Json.tryGetProperty "footers" element
                |> Option.map HeaderFooterTemplates.FromJson
              Elements =
                Json.tryGetArrayProperty "elements" element
                |> Option.map (List.choose DocumentElementTemplate.TryFromJson)
                |> Option.defaultValue [] }

        member ts.Build(ctx: TemplatingContext) =
            let buildElements (ctx: TemplatingContext) =
                ts.Elements
                |> List.fold
                    (fun (elements, ctx) el ->
                        match el with
                        | DocumentElementTemplate.Image -> failwith "TODO"
                        | DocumentElementTemplate.Paragraph p ->
                            match p.Build(ctx) with
                            | TemplatingResult.Single (result, templatingContext) ->
                                elements
                                @ [ Elements.DocumentElement.Paragraph result ],
                                templatingContext
                            | TemplatingResult.Multiple (results, templatingContext) ->
                                elements
                                @ (results
                                   |> List.map Elements.DocumentElement.Paragraph),
                                templatingContext
                        | DocumentElementTemplate.Table t ->
                            match t.build (ctx) with
                            | TemplatingResult.Single (result, newCtx) ->
                                elements
                                @ [ Elements.DocumentElement.Table result ],
                                newCtx
                            | TemplatingResult.Multiple (results, newCtx) ->
                                elements
                                @ (results |> List.map Elements.DocumentElement.Table),
                                newCtx
                        | DocumentElementTemplate.PageBreak -> elements @ [ Elements.DocumentElement.PageBreak ], ctx)
                    ([], ctx)
                |> fst

            let build (ctx) =

                ({ PageSetup = ts.PageSetup
                   Headers = None // TODO
                   Footers = None // TODO
                   Elements = buildElements ctx }: Structure.Section)

            match ts.Scope,
                  ts.ScopeType
                  |> Option.defaultValue SectionScopeType.Standard
                with
            | Some s, SectionScopeType.Standard ->
                ctx.Iterate(s, build)
                |> fun (ctx, els) -> TemplatingResult.Multiple(els, ctx)
            | Some s, SectionScopeType.Inner ->
                ctx.Iterate(s, buildElements)
                |> fun (ctx, els) ->
                    (({ PageSetup = ts.PageSetup
                        Headers = None // TODO
                        Footers = None // TODO
                        Elements = els |> List.concat }: Structure.Section),
                     ctx)
                    |> TemplatingResult.Single

            | None, _ ->
                build ctx
                |> fun s -> TemplatingResult.Single(s, ctx)

    and HeaderFooterTemplates =
        { Primary: HeaderFooterTemplate option
          FirstPage: HeaderFooterTemplate option
          EvenPages: HeaderFooterTemplate option }

        static member FromJson(element: JsonElement) =
            { Primary =
                Json.tryGetProperty "primary" element
                |> Option.map HeaderFooterTemplate.FromJson
              FirstPage =
                Json.tryGetProperty "firstPage" element
                |> Option.map HeaderFooterTemplate.FromJson
              EvenPages =
                Json.tryGetProperty "evenPages" element
                |> Option.map HeaderFooterTemplate.FromJson }

    and HeaderFooterTemplate =
        { Format: Style.ParagraphFormat option
          Style: string option
          Elements: DocumentElementTemplate list }

        static member FromJson(element: JsonElement) =
            { Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Style = Json.tryGetStringProperty "style" element
              Elements =
                Json.tryGetArrayProperty "elements" element
                |> Option.map (List.choose DocumentElementTemplate.TryFromJson)
                |> Option.defaultValue [] }

    and [<RequireQualifiedAccess>] DocumentElementTemplate =
        | Image
        | Paragraph of ParagraphTemplate
        | Table of TableTemplate
        | PageBreak

        static member TryFromJson(element: JsonElement) =
            Json.tryGetStringProperty "type" element
            |> Option.bind (function
                | "image" -> failwith "TODO"
                | "paragraph" ->
                    ParagraphTemplate.FromJson element
                    |> DocumentElementTemplate.Paragraph
                    |> Some
                | "table" ->
                    TableTemplate.FromJson element
                    |> DocumentElementTemplate.Table
                    |> Some
                | "page-break" -> Some DocumentElementTemplate.PageBreak
                | _ -> None)

    and TableTemplate =
        { Scope: Scope option
          Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          TopPadding: Style.Unit option
          BottomPadding: Style.Unit option
          LeftPadding: Style.Unit option
          RightPadding: Style.Unit option
          KeepTogether: bool option
          Columns: TableColumnTemplate list
          Rows: TableRowTemplate list }

        static member FromJson(element: JsonElement) =
            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              Borders =
                Json.tryGetProperty "borders" element
                |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Shading =
                Json.tryGetProperty "shading" element
                |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              TopPadding =
                Json.tryGetProperty "topPadding" element
                |> Option.bind Style.Unit.TryFromJson
              BottomPadding =
                Json.tryGetProperty "bottomPadding" element
                |> Option.bind Style.Unit.TryFromJson
              LeftPadding =
                Json.tryGetProperty "leftPadding" element
                |> Option.bind Style.Unit.TryFromJson
              RightPadding =
                Json.tryGetProperty "rightPadding" element
                |> Option.bind Style.Unit.TryFromJson
              KeepTogether = Json.tryGetBoolProperty "keepTogether" element
              Columns =
                Json.tryGetArrayProperty "columns" element
                |> Option.map (List.map TableColumnTemplate.FromJson)
                |> Option.defaultValue []
              Rows =
                Json.tryGetArrayProperty "rows" element
                |> Option.map (List.map TableRowTemplate.FromJson)
                |> Option.defaultValue [] }

        member tt.build(ctx: TemplatingContext) =
            // TODO handle scope
            let (columns, newCtx) =
                tt.Columns
                |> List.fold
                    (fun (cs, ctx) c ->
                        match c.Build(ctx) with
                        | TemplatingResult.Single (result, newCtx) -> cs @ [ result ], newCtx
                        | TemplatingResult.Multiple (result, newCtx) -> cs @ result, newCtx)
                    ([], ctx)

            let (rows, newCtx2) =
                tt.Rows
                |> List.fold
                    (fun (rs, ctx) r ->
                        match r.Build(ctx) with
                        | TemplatingResult.Single (result, newCtx) -> rs @ [ result ], newCtx
                        | TemplatingResult.Multiple (results, newCtx) -> rs @ results, newCtx)
                    ([], newCtx)

            ({ Borders = tt.Borders
               Format = tt.Format
               Shading = tt.Shading
               Style = tt.Style
               TopPadding = tt.TopPadding
               BottomPadding = tt.BottomPadding
               LeftPadding = tt.LeftPadding
               RightPadding = tt.RightPadding
               KeepTogether = tt.KeepTogether
               Columns = columns
               Rows = rows }: Elements.Table)
            |> fun r -> TemplatingResult.Single(r, newCtx2)

    and TableColumnTemplate =
        { Scope: Scope option
          Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          Width: Style.Unit option
          HeadingFormat: bool option
          KeepWith: int option
          LeftPadding: Style.Unit option
          RightPadding: Style.Unit option }

        static member FromJson(element: JsonElement) =
            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              Borders =
                Json.tryGetProperty "borders" element
                |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Shading =
                Json.tryGetProperty "shading" element
                |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              Width =
                Json.tryGetProperty "width" element
                |> Option.bind Style.Unit.TryFromJson
              HeadingFormat = Json.tryGetBoolProperty "headingFormat" element
              KeepWith = Json.tryGetIntProperty "keepWith" element
              LeftPadding =
                Json.tryGetProperty "leftPadding" element
                |> Option.bind Style.Unit.TryFromJson
              RightPadding =
                Json.tryGetProperty "rightPadding" element
                |> Option.bind Style.Unit.TryFromJson }

        member tct.Build(ctx: TemplatingContext) =
            // TODO handle scope

            ({ Borders = tct.Borders
               Format = tct.Format
               Shading = tct.Shading
               Style = tct.Style
               Width = tct.Width
               HeadingFormat = tct.HeadingFormat
               KeepWith = tct.KeepWith
               LeftPadding = tct.LeftPadding
               RightPadding = tct.RightPadding }: Elements.TableColumn)
            |> fun r -> TemplatingResult.Single(r, ctx)

    and TableRowTemplate =
        { Scope: Scope option
          Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Height: Style.Unit option
          Shading: Style.Shading option
          Style: string option
          TopPadding: Style.Unit option
          BottomPadding: Style.Unit option
          HeadingFormat: bool option
          KeepWith: int option
          VerticalAlignment: Style.VerticalAlignment option
          Cells: TableCellTemplate list }

        static member FromJson(element: JsonElement) =

            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              Borders =
                Json.tryGetProperty "borders" element
                |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Height =
                Json.tryGetProperty "height" element
                |> Option.bind Style.Unit.TryFromJson
              Shading =
                Json.tryGetProperty "shading" element
                |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              TopPadding =
                Json.tryGetProperty "topPadding" element
                |> Option.bind Style.Unit.TryFromJson
              BottomPadding =
                Json.tryGetProperty "bottomPadding" element
                |> Option.bind Style.Unit.TryFromJson
              HeadingFormat = Json.tryGetBoolProperty "headingFormat" element
              KeepWith = Json.tryGetIntProperty "keepWith" element
              VerticalAlignment =
                Json.tryGetIntProperty "verticalAlignment" element
                |> Option.bind Style.VerticalAlignment.Deserialize
              Cells =
                Json.tryGetArrayProperty "cells" element
                |> Option.map (List.map TableCellTemplate.FromJson)
                |> Option.defaultValue [] }

        member trt.Build(ctx: TemplatingContext) =

            let build (ctx: TemplatingContext) =
                let cells =
                    trt.Cells
                    |> List.fold
                        (fun (cs, ctx) tct ->
                            match tct.Build(ctx) with
                            | TemplatingResult.Single (result, newCtx) -> cs @ [ result ], newCtx
                            | TemplatingResult.Multiple (results, newCtx) -> cs @ results, newCtx)
                        ([], ctx)
                    |> fst // Is discarding the new state here wise???

                ({ Borders = trt.Borders
                   Format = trt.Format
                   Height = trt.Height
                   Shading = trt.Shading
                   Style = trt.Style
                   TopPadding = trt.TopPadding
                   BottomPadding = trt.BottomPadding
                   HeadingFormat = trt.HeadingFormat
                   KeepWith = trt.KeepWith
                   VerticalAlignment = trt.VerticalAlignment
                   Cells = cells }: Elements.TableRow)

            match trt.Scope with
            | Some s ->
                ctx.Iterate(s, build)
                |> fun (ctx, rs) -> TemplatingResult.Multiple(rs, ctx)
            | None ->
                build ctx
                |> fun r -> TemplatingResult.Single(r, ctx)

    and TableCellTemplate =
        { Scope: Scope option
          Index: int
          Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          MergeDown: int option
          MergeRight: int option
          VerticalAlignment: Style.VerticalAlignment option
          Elements: CellElementTemplate list }

        static member FromJson(element: JsonElement) =
            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              Index =
                Json.tryGetIntProperty "index" element
                |> Option.defaultValue 0
              Borders =
                Json.tryGetProperty "borders" element
                |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Shading =
                Json.tryGetProperty "shading" element
                |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              MergeDown = Json.tryGetIntProperty "mergeDown" element
              MergeRight = Json.tryGetIntProperty "mergeRight" element
              VerticalAlignment =
                Json.tryGetIntProperty "verticalAlignment" element
                |> Option.bind Style.VerticalAlignment.Deserialize
              Elements =
                Json.tryGetArrayProperty "elements" element
                |> Option.map (fun els -> els |> List.choose CellElementTemplate.TryFromJson)
                |> Option.defaultValue [] }

        member tct.Build(ctx: TemplatingContext) =
            // TODO handle scope

            let (elements, newCtx) =
                tct.Elements
                |> List.fold
                    (fun (els, ctx) el ->
                        match el with
                        | CellElementTemplate.Image -> failwith "Images to be implemented"
                        | CellElementTemplate.Paragraph p ->
                            match p.Build(ctx) with
                            | TemplatingResult.Single (result, templatingContext) ->
                                els @ [ Elements.CellElement.Paragraph result ], templatingContext
                            | TemplatingResult.Multiple (results, templatingContext) ->
                                els
                                @ (results |> List.map Elements.CellElement.Paragraph),
                                templatingContext)
                    ([], ctx)

            ({ Index = tct.Index
               Borders = tct.Borders
               Format = tct.Format
               Shading = tct.Shading
               Style = tct.Style
               MergeDown = tct.MergeDown
               MergeRight = tct.MergeRight
               VerticalAlignment = tct.VerticalAlignment
               Elements = elements }: Elements.TableCell)
            |> fun r -> TemplatingResult.Single(r, newCtx)

    and [<RequireQualifiedAccess>] CellElementTemplate =
        | Image
        | Paragraph of ParagraphTemplate

        static member TryFromJson(element: JsonElement) =
            Json.tryGetStringProperty "type" element
            |> Option.bind (function
                | "image" -> failwith "TODO"
                | "paragraph" ->
                    ParagraphTemplate.FromJson element
                    |> CellElementTemplate.Paragraph
                    |> Some
                | _ -> None)

        member cet.Build(ctx: TemplatingContext) =
            match cet with
            | Image -> failwith "Images to be implemented"
            | Paragraph p -> p.Build(ctx)

    and ParagraphTemplate =
        { Scope: Scope option
          Format: Style.ParagraphFormat option
          Style: string option
          Elements: ParagraphElementTemplate list }

        static member FromJson(element: JsonElement) =
            { Scope =
                Json.tryGetStringProperty "scope" element
                |> Option.bind Scope.TryParse
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Style = Json.tryGetStringProperty "style" element
              Elements =
                Json.tryGetArrayProperty "elements" element
                |> Option.map (fun els ->
                    els
                    |> List.choose ParagraphElementTemplate.TryFromJson)
                |> Option.defaultValue [] }

        member pt.Build(ctx: TemplatingContext) =
            match pt.Scope with
            | Some s ->
                let build (tc: TemplatingContext) =
                    ({ Format = pt.Format
                       Style = pt.Style
                       Elements = pt.Elements |> List.map (fun el -> el.Build(tc)) }: Elements.Paragraph)

                ctx.Iterate(s, build)
                |> fun (ctx, els) -> TemplatingResult.Multiple(els, ctx)
            | None ->
                // No scope, so simply build the paragraph
                ({ Format = pt.Format
                   Style = pt.Style
                   Elements = pt.Elements |> List.map (fun el -> el.Build(ctx)) }: Elements.Paragraph)
                |> fun r -> TemplatingResult.Single(r, ctx)

    and [<RequireQualifiedAccess>] ParagraphElementTemplate =
        | Image
        | Space
        | Tab
        | Text of TextTemplate
        | FormattedText of FormattedTextTemplate
        | LineBreak
        | Conditional of ParagraphElementTemplate

        static member TryFromJson(element: JsonElement) =
            Json.tryGetStringProperty "type" element
            |> Option.bind (function
                | "image" -> failwith "TODO"
                | "space" -> Some ParagraphElementTemplate.Space
                | "tab" -> Some ParagraphElementTemplate.Tab
                | "text" ->
                    TextTemplate.FromJson element
                    |> Option.map ParagraphElementTemplate.Text
                | "formatted-text" -> None
                | "line-break" -> None
                | "conditional" -> failwith "TODO"
                | _ -> None)

        member pet.Build(ctx: TemplatingContext) =
            match pet with
            | Image -> Elements.ParagraphElement.Image
            | Space -> Elements.ParagraphElement.Space
            | Tab -> Elements.ParagraphElement.Tab
            | Text tt -> tt.Build(ctx) |> Elements.ParagraphElement.Text
            | FormattedText ftt ->
                ftt.Build(ctx)
                |> Elements.ParagraphElement.FormattedText
            | LineBreak -> Elements.ParagraphElement.LineBreak
            | Conditional _ -> failwith "To implement"

    and TextTemplate =
        { Content: Content list }

        static member FromJson(element: JsonElement) =
            Json.tryGetStringProperty "content" element
            |> Option.map (fun c ->
                // parse the string content
                let content =
                    Internal.contentRegex.Matches c
                    |> List.ofSeq
                    |> List.map (fun m ->
                        let (group, value) =
                            m.Groups
                            |> Seq.skip 1
                            |> Seq.tryFind (fun g -> g.Success)
                            |> Option.map (fun g ->g.Name, g.Value)
                            |> Option.defaultValue ("literal", m.Value)
                        match group with
                        | "value" ->
                            Data.Path.Parse value
                            |> Content.Value
                        | "preprocessor" ->
                            Content.Processor (*value*)
                        | "literal"
                        | _ -> Content.Literal value)
                { Content = content })

        member tt.Build(ctx: TemplatingContext) =
            tt.Content
            |> List.fold
                (fun acc c ->
                    match c with
                    | Content.Literal v -> Some v :: acc
                    | Content.Value v ->
                        (ctx.Data.Iterators.ExpandPath v
                         |> ctx.ResolveValue)
                        :: acc
                    | Content.Processor -> failwith "TODO")
                []
            |> List.choose id
            |> List.rev
            |> fun r -> ({ Content = r |> String.concat "" }: Elements.Text)

    and FormattedTextTemplate =
        { Bold: bool option
          Color: Style.Color option
          Italic: bool option
          Size: Style.Unit option
          Subscript: bool option
          Superscript: bool option
          Underline: Style.Underline option
          Font: Style.Font option
          Style: string option
          Elements: ParagraphElementTemplate list }

        static member FromJson(element: JsonElement) =
            { Bold = Json.tryGetBoolProperty "bold" element
              Color =
                Json.tryGetProperty "color" element
                |> Option.bind Style.Color.TryFromJson
              Italic = Json.tryGetBoolProperty "italic" element
              Size =
                Json.tryGetProperty "size" element
                |> Option.bind Style.Unit.TryFromJson
              Subscript = Json.tryGetBoolProperty "subscript" element
              Superscript = Json.tryGetBoolProperty "superscript" element
              Underline =
                Json.tryGetIntProperty "underline" element
                |> Option.bind Style.Underline.Deserialize
              Font =
                Json.tryGetProperty "font" element
                |> Option.map Style.Font.FromJson
              Style = Json.tryGetStringProperty "style" element
              Elements =
                Json.tryGetArrayProperty "elements" element
                |> Option.map (fun els ->
                    els
                    |> List.choose ParagraphElementTemplate.TryFromJson)
                |> Option.defaultValue [] }

        member ftt.Build(ctx: TemplatingContext) =
            ({ Bold = ftt.Bold
               Color = ftt.Color
               Italic = ftt.Italic
               Size = ftt.Size
               Subscript = ftt.Subscript
               Superscript = ftt.Subscript
               Underline = ftt.Underline
               Font = ftt.Font
               Style = ftt.Style
               Elements = ftt.Elements |> List.map (fun el -> el.Build(ctx)) }: Elements.FormattedText)
