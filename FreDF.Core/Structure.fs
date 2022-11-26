namespace FreDF.Core

[<RequireQualifiedAccess>]
module Structure =

    open System.Text.Json
    open ToolBox.Core

    type PdfDocument =
        { Sections: Section list }

        member pdf.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Document()

            pdf.Sections
            |> List.iter (fun s -> obj.Add(section = s.ToDocObj()))

            obj

    and Section =
        { PageSetup: PageSetup option
          Headers: HeaderFooters option
          Footers: HeaderFooters option
          Elements: Elements.DocumentElement list }

        member s.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Section()

            s.PageSetup
            |> Option.iter (fun ps -> obj.PageSetup <- ps.toDocObj ())

            s.Headers
            |> Option.iter (fun h -> obj.Headers <- h.ToDocObj())

            s.Footers
            |> Option.iter (fun f -> obj.Footers <- f.ToDocObj())

            s.Elements
            |> List.iter (function
                | Elements.DocumentElement.Paragraph p -> obj.Add(paragraph = p.ToDocObj())
                | Elements.DocumentElement.Table t -> obj.Add(table = t.ToDocObj())
                | _ -> ())

            obj

    and HeaderFooters =
        { Primary: HeaderFooter option
          FirstPage: HeaderFooter option
          EvenPages: HeaderFooter option }

        member hf.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.HeadersFooters()

            hf.Primary
            |> Option.iter (fun p -> obj.Primary <- p.ToDocObj())

            hf.FirstPage
            |> Option.iter (fun fp -> obj.FirstPage <- fp.ToDocObj())

            hf.EvenPages
            |> Option.iter (fun ep -> obj.EvenPage <- ep.ToDocObj())

            obj

    and HeaderFooter =
        { Format: Style.ParagraphFormat option
          Style: string option
          Elements: Elements.DocumentElement list }

        member hf.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.HeaderFooter()

            hf.Format
            |> Option.iter (fun f -> obj.Format <- f.ToDocObj())

            hf.Style |> Option.iter (fun s -> obj.Style <- s)

            hf.Elements
            |> List.iter (function
                | Elements.DocumentElement.Image -> ()
                | Elements.DocumentElement.Paragraph p -> obj.Elements.Add(p.ToDocObj())
                | Elements.DocumentElement.Table t -> obj.Elements.Add(t.ToDocObj())
                | Elements.DocumentElement.PageBreak ->
                    //NOTE page breaks not supported in headers and footers
                    ())

            obj

    and PageSetup =
        { Orientation: PageOrientation option
          TopMargin: Style.Unit option
          BottomMargin: Style.Unit option
          LeftMargin: Style.Unit option
          RightMargin: Style.Unit option
          HeaderDistance: Style.Unit option
          FooterDistance: Style.Unit option
          MirrorMargins: bool option
          PageHeight: Style.Unit option
          PageWidth: Style.Unit option
          PageFormat: PageFormat option
          SectionStart: BreakType option
          StartingNumber: int option
          HorizontalPageBreak: bool option
          DifferentFirstPageHeaderFooter: bool option
          OddAndEvenPagesHeaderFooter: bool option }

        static member Blank() =
            { Orientation = None
              TopMargin = None
              BottomMargin = None
              LeftMargin = None
              RightMargin = None
              HeaderDistance = None
              FooterDistance = None
              MirrorMargins = None
              PageHeight = None
              PageWidth = None
              PageFormat = None
              SectionStart = None
              StartingNumber = None
              HorizontalPageBreak = None
              DifferentFirstPageHeaderFooter = None
              OddAndEvenPagesHeaderFooter = None }

        static member Default() =
            { Orientation = Some PageOrientation.Portrait
              TopMargin = None
              BottomMargin = None
              LeftMargin = None
              RightMargin = None
              HeaderDistance = None
              FooterDistance = None
              MirrorMargins = None
              PageHeight = None
              PageWidth = None
              PageFormat = Some PageFormat.A4
              SectionStart = None
              StartingNumber = None
              HorizontalPageBreak = None
              DifferentFirstPageHeaderFooter = None
              OddAndEvenPagesHeaderFooter = None }

        static member FromJson(element: JsonElement) =
            { Orientation =
                Json.tryGetIntProperty "orientation" element
                |> Option.bind PageOrientation.TryDeserialize
              TopMargin =
                Json.tryGetProperty "topMargin" element
                |> Option.bind Style.Unit.TryFromJson
              BottomMargin =
                Json.tryGetProperty "bottomMargin" element
                |> Option.bind Style.Unit.TryFromJson
              LeftMargin =
                Json.tryGetProperty "leftMargin" element
                |> Option.bind Style.Unit.TryFromJson
              RightMargin =
                Json.tryGetProperty "rightMargin" element
                |> Option.bind Style.Unit.TryFromJson
              HeaderDistance =
                Json.tryGetProperty "headerDistance" element
                |> Option.bind Style.Unit.TryFromJson
              FooterDistance =
                Json.tryGetProperty "footerDistance" element
                |> Option.bind Style.Unit.TryFromJson
              MirrorMargins = Json.tryGetBoolProperty "mirrorMargins" element
              PageHeight =
                Json.tryGetProperty "pageHeight" element
                |> Option.bind Style.Unit.TryFromJson
              PageWidth =
                Json.tryGetProperty "pageWidth" element
                |> Option.bind Style.Unit.TryFromJson
              PageFormat =
                Json.tryGetIntProperty "pageFormat" element
                |> Option.bind PageFormat.TryDeserialize
              SectionStart =
                Json.tryGetIntProperty "sectionStart" element
                |> Option.bind BreakType.TryDeserialize
              StartingNumber = Json.tryGetIntProperty "startingNumber" element
              HorizontalPageBreak = Json.tryGetBoolProperty "horizontalPageBreak" element
              DifferentFirstPageHeaderFooter = Json.tryGetBoolProperty "differentFirstPageHeaderFooter" element
              OddAndEvenPagesHeaderFooter = Json.tryGetBoolProperty "oddAndEvenPagesHeaderFooter" element }

        member ps.toDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.PageSetup()

            ps.Orientation
            |> Option.iter (fun po -> obj.Orientation <- po.ToDocObj())

            ps.TopMargin
            |> Option.iter (fun tm -> obj.TopMargin <- tm.ToDocObj())

            ps.BottomMargin
            |> Option.iter (fun bm -> obj.BottomMargin <- bm.ToDocObj())

            ps.LeftMargin
            |> Option.iter (fun lm -> obj.LeftMargin <- lm.ToDocObj())

            ps.RightMargin
            |> Option.iter (fun rm -> obj.RightMargin <- rm.ToDocObj())

            ps.HeaderDistance
            |> Option.iter (fun hd -> obj.HeaderDistance <- hd.ToDocObj())

            ps.FooterDistance
            |> Option.iter (fun fd -> obj.FooterDistance <- fd.ToDocObj())

            ps.MirrorMargins
            |> Option.iter (fun mm -> obj.MirrorMargins <- mm)

            ps.PageHeight
            |> Option.iter (fun ph -> obj.PageHeight <- ph.ToDocObj())

            ps.PageWidth
            |> Option.iter (fun pw -> obj.PageWidth <- pw.ToDocObj())

            ps.PageFormat
            |> Option.iter (fun pf -> obj.PageFormat <- pf.ToDocObj())

            ps.SectionStart
            |> Option.iter (fun s -> obj.SectionStart <- s.ToDocObj())

            ps.StartingNumber
            |> Option.iter (fun sn -> obj.StartingNumber <- sn)

            ps.HorizontalPageBreak
            |> Option.iter (fun hpb -> obj.HorizontalPageBreak <- hpb)

            ps.DifferentFirstPageHeaderFooter
            |> Option.iter (fun dfphf -> obj.DifferentFirstPageHeaderFooter <- dfphf)

            ps.OddAndEvenPagesHeaderFooter
            |> Option.iter (fun oephf -> obj.OddAndEvenPagesHeaderFooter <- oephf)

            obj

    and [<RequireQualifiedAccess>] PageOrientation =
        | Portrait
        | Landscape

        static member TryDeserialize(value: int) =
            match value with
            | 0 -> Some PageOrientation.Portrait
            | 1 -> Some PageOrientation.Landscape
            | _ -> None

        static member Deserialize(value: int) =
            PageOrientation.TryDeserialize value
            |> Option.defaultValue PageOrientation.Portrait

        member po.ToDocObj() =
            match po with
            | PageOrientation.Portrait -> MigraDocCore.DocumentObjectModel.Orientation.Portrait
            | PageOrientation.Landscape -> MigraDocCore.DocumentObjectModel.Orientation.Landscape

    and [<RequireQualifiedAccess>] PageFormat =
        | A0
        | A1
        | A2
        | A3
        | A4
        | A5
        | A6
        | B5
        | Letter
        | Legal
        | Ledger
        | P11x17

        static member TryDeserialize(value: int) =
            match value with
            | 0 -> Some A0
            | 1 -> Some A1
            | 2 -> Some A2
            | 3 -> Some A3
            | 4 -> Some A4
            | 5 -> Some A5
            | 6 -> Some A6
            | 7 -> Some B5
            | 8 -> Some Letter
            | 9 -> Some Legal
            | 10 -> Some Ledger
            | 11 -> Some P11x17
            | _ -> None

        static member Deserialize(value: int) =
            PageFormat.TryDeserialize value
            |> Option.defaultValue PageFormat.A4

        member pf.ToDocObj() =
            match pf with
            | PageFormat.A0 -> MigraDocCore.DocumentObjectModel.PageFormat.A0
            | PageFormat.A1 -> MigraDocCore.DocumentObjectModel.PageFormat.A1
            | PageFormat.A2 -> MigraDocCore.DocumentObjectModel.PageFormat.A2
            | PageFormat.A3 -> MigraDocCore.DocumentObjectModel.PageFormat.A3
            | PageFormat.A4 -> MigraDocCore.DocumentObjectModel.PageFormat.A4
            | PageFormat.A5 -> MigraDocCore.DocumentObjectModel.PageFormat.A5
            | PageFormat.A6 -> MigraDocCore.DocumentObjectModel.PageFormat.A6
            | PageFormat.B5 -> MigraDocCore.DocumentObjectModel.PageFormat.A5
            | PageFormat.Letter -> MigraDocCore.DocumentObjectModel.PageFormat.Letter
            | PageFormat.Legal -> MigraDocCore.DocumentObjectModel.PageFormat.Legal
            | PageFormat.Ledger -> MigraDocCore.DocumentObjectModel.PageFormat.Ledger
            | PageFormat.P11x17 -> MigraDocCore.DocumentObjectModel.PageFormat.P11x17

    and [<RequireQualifiedAccess>] BreakType =
        | BreakNextPage
        | BreakEvenPage
        | BreakOddPage

        static member TryDeserialize(value: int) =
            match value with
            | 0 -> Some BreakType.BreakNextPage
            | 1 -> Some BreakType.BreakEvenPage
            | 2 -> Some BreakType.BreakOddPage
            | _ -> None

        member bt.ToDocObj() =
            match bt with
            | BreakType.BreakNextPage -> MigraDocCore.DocumentObjectModel.BreakType.BreakNextPage
            | BreakType.BreakEvenPage -> MigraDocCore.DocumentObjectModel.BreakType.BreakEvenPage
            | BreakType.BreakOddPage -> MigraDocCore.DocumentObjectModel.BreakType.BreakOddPage
