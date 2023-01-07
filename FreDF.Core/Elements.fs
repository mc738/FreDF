namespace FreDF.Core


[<RequireQualifiedAccess>]
module Elements =

    open System.Text.Json
    open ToolBox.Core

    // All section elements
    // * Chart
    // * Image *
    // * Paragraph *
    // * Table *
    // * Page break *
    // * Text frame
    [<RequireQualifiedAccess>]
    type DocumentElement =
        | Image of Image
        | Paragraph of Paragraph
        | Table of Table
        | PageBreak

    // All cell elements
    // * Chart
    // * Image *
    // * Paragraph *
    // * Text frame
    and [<RequireQualifiedAccess>] CellElement =
        | Image of Image
        | Paragraph of Paragraph

    // All paragraph elements
    // * Bookmark
    // * Char
    // * Character
    // * Footnote
    // * Hyper link
    // * Image *
    // * Space *
    // * Tab *
    // * Text *
    // * DateField
    // * Formatted text *
    // * Info field
    // * Line break *
    // * Page field
    // * Section field
    // * Num pages field
    // * Page ref field
    // * Section pages field

    and [<RequireQualifiedAccess>] ParagraphElement =
        | Image of Image
        | Space
        | Tab
        | Text of Text
        | FormattedText of FormattedText
        | LineBreak

    and Table =
        { Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          TopPadding: Style.Unit option
          BottomPadding: Style.Unit option
          LeftPadding: Style.Unit option
          RightPadding: Style.Unit option
          KeepTogether: bool option
          Columns: TableColumn list
          Rows: TableRow list }

        static member FromJson(element: JsonElement) =
            { Borders = Json.tryGetProperty "borders" element |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Shading = Json.tryGetProperty "shading" element |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              TopPadding = Json.tryGetProperty "topPadding" element |> Option.bind Style.Unit.TryFromJson
              BottomPadding =
                Json.tryGetProperty "bottomPadding" element
                |> Option.bind Style.Unit.TryFromJson
              LeftPadding = Json.tryGetProperty "leftPadding" element |> Option.bind Style.Unit.TryFromJson
              RightPadding = Json.tryGetProperty "rightPadding" element |> Option.bind Style.Unit.TryFromJson
              KeepTogether = Json.tryGetBoolProperty "keepTogether" element
              Columns =
                Json.tryGetArrayProperty "columns" element
                |> Option.map (List.map TableColumn.FromJson)
                |> Option.defaultValue []
              Rows =
                Json.tryGetArrayProperty "rows" element
                |> Option.map (List.map TableRow.FromJson)
                |> Option.defaultValue [] }

        member t.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.Tables.Table()

            t.Borders |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            t.Format |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            t.Shading |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            t.Style |> Option.iter (fun s -> obj.Style <- s)

            t.TopPadding |> Option.iter (fun tp -> obj.TopPadding <- tp.ToDocObj())

            t.BottomPadding |> Option.iter (fun bp -> obj.BottomPadding <- bp.ToDocObj())

            t.LeftPadding |> Option.iter (fun lp -> obj.LeftPadding <- lp.ToDocObj())

            t.RightPadding |> Option.iter (fun rp -> obj.RightPadding <- rp.ToDocObj())

            t.KeepTogether |> Option.iter (fun kt -> obj.KeepTogether <- kt)

            let columns = MigraDocCore.DocumentObjectModel.Tables.Columns()

            let rows = MigraDocCore.DocumentObjectModel.Tables.Rows()

            t.Columns |> List.iter (fun c -> c.ToDocObj() |> columns.Add)

            t.Rows |> List.iter (fun r -> r.ToDocObj() |> rows.Add)

            obj.Columns <- columns
            obj.Rows <- rows

            obj

    and TableColumn =
        { Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          Width: Style.Unit option
          HeadingFormat: bool option
          KeepWith: int option
          LeftPadding: Style.Unit option
          RightPadding: Style.Unit option }

        static member FromJson(element: JsonElement) =
            { Borders = Json.tryGetProperty "borders" element |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Shading = Json.tryGetProperty "shading" element |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              Width = Json.tryGetProperty "width" element |> Option.bind Style.Unit.TryFromJson
              HeadingFormat = Json.tryGetBoolProperty "headingFormat" element
              KeepWith = Json.tryGetIntProperty "keepWith" element
              LeftPadding = Json.tryGetProperty "leftPadding" element |> Option.bind Style.Unit.TryFromJson
              RightPadding = Json.tryGetProperty "rightPadding" element |> Option.bind Style.Unit.TryFromJson }

        member tc.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.Tables.Column()

            tc.Borders |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            tc.Format |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            tc.Shading |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            tc.Style |> Option.iter (fun s -> obj.Style <- s)

            tc.Width |> Option.iter (fun w -> obj.Width <- w.ToDocObj())

            tc.HeadingFormat |> Option.iter (fun hf -> obj.HeadingFormat <- hf)

            tc.KeepWith |> Option.iter (fun kw -> obj.KeepWith <- kw)

            tc.LeftPadding |> Option.iter (fun lp -> obj.LeftPadding <- lp.ToDocObj())

            tc.RightPadding |> Option.iter (fun rp -> obj.RightPadding <- rp.ToDocObj())


            obj

    and TableRow =
        { Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Height: Style.Unit option
          Shading: Style.Shading option
          Style: string option
          TopPadding: Style.Unit option
          BottomPadding: Style.Unit option
          HeadingFormat: bool option
          KeepWith: int option
          VerticalAlignment: Style.VerticalAlignment option
          Cells: TableCell list }

        static member FromJson(element: JsonElement) =
            { Borders = Json.tryGetProperty "borders" element |> Option.map Style.Borders.FromJson
              Format =
                Json.tryGetProperty "format" element
                |> Option.map Style.ParagraphFormat.FromJson
              Height = Json.tryGetProperty "height" element |> Option.bind Style.Unit.TryFromJson
              Shading = Json.tryGetProperty "shading" element |> Option.bind Style.Shading.TryFromJson
              Style = Json.tryGetStringProperty "style" element
              TopPadding = Json.tryGetProperty "topPadding" element |> Option.bind Style.Unit.TryFromJson
              BottomPadding =
                Json.tryGetProperty "bottomPadding" element
                |> Option.bind Style.Unit.TryFromJson
              HeadingFormat = Json.tryGetBoolProperty "headingFormat" element
              KeepWith = Json.tryGetIntProperty "keepWith" element
              VerticalAlignment =
                Json.tryGetIntProperty "verticalAlignment" element
                |> Option.bind Style.VerticalAlignment.Deserialize
              Cells = [] }

        member tr.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.Tables.Row()

            tr.Borders |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            tr.Format |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            tr.Height |> Option.iter (fun h -> obj.Height <- h.ToDocObj())

            tr.Shading |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            tr.Style |> Option.iter (fun s -> obj.Style <- s)

            tr.TopPadding |> Option.iter (fun tp -> obj.TopPadding <- tp.ToDocObj())

            tr.BottomPadding |> Option.iter (fun bp -> obj.BottomPadding <- bp.ToDocObj())

            tr.HeadingFormat |> Option.iter (fun hf -> obj.HeadingFormat <- hf)

            tr.KeepWith |> Option.iter (fun kw -> obj.KeepWith <- kw)

            tr.VerticalAlignment
            |> Option.iter (fun va -> obj.VerticalAlignment <- va.ToDocObj())

            tr.Cells
            |> List.iter (fun c ->
                let cell = obj.Cells.[c.Index]

                c.Borders |> Option.iter (fun b -> cell.Borders <- b.ToDocObj())

                c.Format |> Option.iter (fun f -> cell.Format <- f.ToDocObj())

                c.Shading |> Option.iter (fun s -> cell.Shading <- s.ToDocObj())

                c.Style |> Option.iter (fun s -> cell.Style <- s)

                c.MergeDown |> Option.iter (fun md -> cell.MergeDown <- md)

                c.MergeRight |> Option.iter (fun mr -> cell.MergeRight <- mr)

                c.VerticalAlignment
                |> Option.iter (fun va -> cell.VerticalAlignment <- va.ToDocObj())

                c.Elements
                |> List.iter (function
                    | CellElement.Paragraph p -> cell.Add(paragraph = p.ToDocObj())
                    | _ -> ()))

            obj

    and TableCell =
        { Index: int
          Borders: Style.Borders option
          Format: Style.ParagraphFormat option
          Shading: Style.Shading option
          Style: string option
          MergeDown: int option
          MergeRight: int option
          VerticalAlignment: Style.VerticalAlignment option
          Elements: CellElement list }

    and Paragraph =
        { Elements: ParagraphElement list
          Format: Style.ParagraphFormat option
          Style: string option }

        static member Default() =
            ({ Index = 0
               Borders = None
               Format = None
               Shading = None
               Style = None
               MergeDown = None
               MergeRight = None
               VerticalAlignment = None
               Elements = [] }: TableCell)

        member p.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.Paragraph()

            p.Format |> Option.iter (fun f -> obj.Format <- f.ToDocObj())
            p.Style |> Option.iter (fun s -> obj.Style <- s)

            p.Elements
            |> List.iter (function
                | ParagraphElement.Image i -> obj.Add(image = i.ToDocObj())
                | ParagraphElement.Space -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Blank)
                | ParagraphElement.Tab -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Tab)
                | ParagraphElement.Text t -> obj.Add(text = t.ToDocObj())
                | ParagraphElement.FormattedText ft -> obj.Add(formattedText = ft.ToDocObj())
                | ParagraphElement.LineBreak -> obj.AddLineBreak())

            obj

    and Text =
        { Content: string }

        member t.ToDocObj() =
            MigraDocCore.DocumentObjectModel.Text(t.Content)

    and FormattedText =
        { Bold: bool option
          Color: Style.Color option
          Italic: bool option
          Size: Style.Unit option
          Subscript: bool option
          Superscript: bool option
          Underline: Style.Underline option
          Font: Style.Font option
          Style: string option
          Elements: ParagraphElement list }

        member ft.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.FormattedText()

            ft.Bold |> Option.iter (fun v -> obj.Bold <- v)

            ft.Color |> Option.iter (fun c -> obj.Color <- c.ToDocObj())

            ft.Italic |> Option.iter (fun i -> obj.Italic <- i)

            ft.Font |> Option.iter (fun f -> obj.Font <- f.ToDocObj())

            ft.Size |> Option.iter (fun s -> obj.Size <- s.ToDocObj())

            ft.Style |> Option.iter (fun s -> obj.Style <- s)

            ft.Subscript |> Option.iter (fun s -> obj.Subscript <- s)

            ft.Superscript |> Option.iter (fun s -> obj.Superscript <- s)

            ft.Underline |> Option.iter (fun ul -> obj.Underline <- ul.ToDocObj())

            ft.Elements
            |> List.iter (function
                | ParagraphElement.Image img -> obj.Add(image = img.ToDocObj())
                | ParagraphElement.Space -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Blank)
                | ParagraphElement.Tab -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Tab)
                | ParagraphElement.Text t -> obj.Add(text = t.ToDocObj())
                | ParagraphElement.FormattedText ft -> obj.Add(formattedText = ft.ToDocObj())
                | ParagraphElement.LineBreak -> obj.AddLineBreak())

            obj
            
    and Image =
        { Source: string
          Height: Style.Unit option
          Width: Style.Unit option
          Left: float option
          Top: float option
          Resolution: float option
          // FillFormat
          // LineFormat
          // Picture format
          // RelativeHorizontal
          // RelativeVertical
          ScaleHeight: float option
          ScaleWidth: float option
          // WrapFormat
          LockAspectRatio: bool option }

        member i.ToDocObj() =
            let obj = MigraDocCore.DocumentObjectModel.Shapes.Image()

            // TODO support different image source types
            obj.Source <-
                MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes.ImageSource.FromFile(i.Source)

            i.Height |> Option.iter (fun h -> obj.Height <- h.ToDocObj())
            i.Width |> Option.iter (fun w -> obj.Width <- w.ToDocObj())
            i.Left |> Option.iter (fun l -> obj.Left <- l)
            i.Top |> Option.iter (fun t -> obj.Top <- t)
            i.Resolution |> Option.iter (fun r -> obj.Resolution <- r)
            i.ScaleHeight |> Option.iter (fun sh -> obj.ScaleHeight <- sh)
            i.ScaleWidth |> Option.iter (fun sw -> obj.ScaleWidth <- sw)
            i.LockAspectRatio |> Option.iter (fun lar -> obj.LockAspectRatio <- lar)
            obj
