﻿namespace FreDF.Core

[<RequireQualifiedAccess>]
module Elements =

    // All section elements
    // * Chart
    // * Image *
    // * Paragraph *
    // * Table *
    // * Page break *
    // * Text frame
    [<RequireQualifiedAccess>]
    type DocumentElement =
        | Image
        | Paragraph of Paragraph
        | Table of TableDefinition
        | PageBreak

    // All cell elements
    // * Chart
    // * Image *
    // * Paragraph *
    // * Text frame
    and [<RequireQualifiedAccess>] CellElement =
        | Image
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
        | Image
        | Space
        | Tab
        | Text of Text
        | FormattedText of FormattedText
        | LineBreak


    and TableDefinition =
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

        member t.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Tables.Table()

            t.Borders
            |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            t.Format
            |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            t.Shading
            |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            t.Style |> Option.iter (fun s -> obj.Style <- s)

            t.TopPadding
            |> Option.iter (fun tp -> obj.TopPadding <- tp.ToDocObj())

            t.BottomPadding
            |> Option.iter (fun bp -> obj.BottomPadding <- bp.ToDocObj())

            t.LeftPadding
            |> Option.iter (fun lp -> obj.LeftPadding <- lp.ToDocObj())

            t.RightPadding
            |> Option.iter (fun rp -> obj.RightPadding <- rp.ToDocObj())

            t.KeepTogether
            |> Option.iter (fun kt -> obj.KeepTogether <- kt)

            let columns =
                MigraDocCore.DocumentObjectModel.Tables.Columns()

            let rows =
                MigraDocCore.DocumentObjectModel.Tables.Rows()

            t.Columns
            |> List.iter (fun c -> c.ToDocObj() |> columns.Add)

            t.Rows
            |> List.iter (fun r -> r.ToDocObj() |> rows.Add)

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

        member tc.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Tables.Column()

            tc.Borders
            |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            tc.Format
            |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            tc.Shading
            |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            tc.Style |> Option.iter (fun s -> obj.Style <- s)

            tc.Width
            |> Option.iter (fun w -> obj.Width <- w.ToDocObj())

            tc.HeadingFormat
            |> Option.iter (fun hf -> obj.HeadingFormat <- hf)

            tc.KeepWith
            |> Option.iter (fun kw -> obj.KeepWith <- kw)

            tc.LeftPadding
            |> Option.iter (fun lp -> obj.LeftPadding <- lp.ToDocObj())

            tc.RightPadding
            |> Option.iter (fun rp -> obj.RightPadding <- rp.ToDocObj())


            obj

    and TableRow =
        { Borders: Style.Borders option
          Cells: TableCell list
          Format: Style.ParagraphFormat option
          Height: Style.Unit option
          Shading: Style.Shading option
          Style: string option
          TopPadding: Style.Unit option
          BottomPadding: Style.Unit option
          HeadingFormat: bool option
          KeepWith: int option
          VerticalAlignment: Style.VerticalAlignment option }

        member tr.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Tables.Row()

            tr.Borders
            |> Option.iter (fun b -> obj.Borders <- b.ToDocObj())

            tr.Format
            |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            tr.Height
            |> Option.iter (fun h -> obj.Height <- h.ToDocObj())

            tr.Shading
            |> Option.iter (fun s -> obj.Shading <- s.ToDocObj())

            tr.Style |> Option.iter (fun s -> obj.Style <- s)

            tr.TopPadding
            |> Option.iter (fun tp -> obj.TopPadding <- tp.ToDocObj())

            tr.BottomPadding
            |> Option.iter (fun bp -> obj.BottomPadding <- bp.ToDocObj())

            tr.HeadingFormat
            |> Option.iter (fun hf -> obj.HeadingFormat <- hf)

            tr.KeepWith
            |> Option.iter (fun kw -> obj.KeepWith <- kw)

            tr.VerticalAlignment
            |> Option.iter (fun va -> obj.VerticalAlignment <- va.ToDocObj())

            tr.Cells
            |> List.iter (fun c ->
                let cell = obj.Cells.[c.Index]

                c.Borders
                |> Option.iter (fun b -> cell.Borders <- b.ToDocObj())

                c.Format
                |> Option.iter (fun f -> cell.Format <- f.ToDocObj())

                c.Shading
                |> Option.iter (fun s -> cell.Shading <- s.ToDocObj())

                c.Style |> Option.iter (fun s -> cell.Style <- s)

                c.MergeDown
                |> Option.iter (fun md -> cell.MergeDown <- md)

                c.MergeRight
                |> Option.iter (fun mr -> cell.MergeRight <- mr)

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
            let obj =
                MigraDocCore.DocumentObjectModel.Paragraph()

            p.Elements
            |> List.iter (function
                | ParagraphElement.Image -> ()
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
            let obj =
                MigraDocCore.DocumentObjectModel.FormattedText()

            ft.Bold |> Option.iter (fun v -> obj.Bold <- v)

            ft.Color
            |> Option.iter (fun c -> obj.Color <- c.ToDocObj())

            ft.Italic
            |> Option.iter (fun i -> obj.Italic <- i)

            ft.Font
            |> Option.iter (fun f -> obj.Font <- f.ToDocObj())

            ft.Size
            |> Option.iter (fun s -> obj.Size <- s.ToDocObj())

            ft.Style |> Option.iter (fun s -> obj.Style <- s)

            ft.Subscript
            |> Option.iter (fun s -> obj.Subscript <- s)

            ft.Superscript
            |> Option.iter (fun s -> obj.Superscript <- s)

            ft.Underline
            |> Option.iter (fun ul -> obj.Underline <- ul.ToDocObj())

            ft.Elements
            |> List.iter (function
                | ParagraphElement.Image -> ()
                | ParagraphElement.Space -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Blank)
                | ParagraphElement.Tab -> obj.Add(MigraDocCore.DocumentObjectModel.Character.Tab)
                | ParagraphElement.Text t -> obj.Add(text = t.ToDocObj())
                | ParagraphElement.FormattedText ft -> obj.Add(formattedText = ft.ToDocObj())
                | ParagraphElement.LineBreak -> obj.AddLineBreak())

            obj