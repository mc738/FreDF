namespace FreDF.Core


module Templating =

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
    and Scope = { Name: string; Path: Data.Path }

    and SectionScopeType =
        | Standard
        | Inner

    and TemplatingResult<'T> =
        | Single of Result: 'T * Context: TemplatingContext
        | Multiple of Result: 'T list * Context: TemplatingContext

    type ConditionType = Bespoke of bool

    type Template =
        { Sections: TemplateSection list }

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

    and HeaderFooterTemplate =
        { Format: Style.ParagraphFormat option
          Style: string option
          Elements: Elements.DocumentElement list }

    and [<RequireQualifiedAccess>] DocumentElementTemplate =
        | Image
        | Paragraph of ParagraphTemplate
        | Table of TableTemplate
        | PageBreak

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

        member cet.Build(ctx: TemplatingContext) =
            match cet with
            | Image -> failwith "Images to be implemented"
            | Paragraph p -> p.Build(ctx)

    and ParagraphTemplate =
        { Scope: Scope option
          Format: Style.ParagraphFormat option
          Style: string option
          Elements: ParagraphElementTemplate list }

        member pt.Build(ctx: TemplatingContext) =
            match pt.Scope with
            | Some s ->
                // Has a scope so...
                // 1. Open the scope.
                let build (tc: TemplatingContext) =
                    //match tc.Data

                    ({ Format = pt.Format
                       Style = pt.Style
                       Elements = pt.Elements |> List.map (fun el -> el.Build(tc)) }: Elements.Paragraph)

                ctx.Iterate(s, build)
                |> fun (ctx, els) -> TemplatingResult.Multiple(els, ctx)

            //newState.Iterate(s.Name, fun tc -> None)

            // 2. While the scope returns values build.

            // Generate a path for each value?
            // Some quote-items becomes
            // (0)__quote-items
            // (1)__quote-items
            // (2)__quote-items
            // ??
            //
            // OR
            //
            // resolve path with value, handle, and again until value is 0?


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

        member tt.Build(ctx: TemplatingContext) =
            tt.Content
            |> List.fold
                (fun acc c ->
                    match c with
                    | Content.Literal v -> Some v :: acc
                    | Content.Value v ->
                        (ctx.Data.Iterators.ExpandPath v
                         |> ctx.ResolveValue)
                        :: acc)
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
