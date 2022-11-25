namespace FreDF.Core

[<RequireQualifiedAccess>]
module Style =

    open System.Text.Json
    open ToolBox.Core

    // NOTE - styles are for internal use.
    // The naming/types can be confusing compared with MigraDocCore.DocumentObjectModel classes.
    // Each one has a method .ToDocObj() which converts to a MigraDocCore.DocumentObjectModel object for use in generating PDFs.
    // This module is made for serialization and deserialize, that is way MigraDocCore.DocumentObjectModel is fully qualified so much.
    // Also due to how PdfSharp/MigraDocCore works there is a lot of internal mutability here.

    [<RequireQualifiedAccess>]
    type Unit =
        | Centimeter of Value: float
        | Point of Value: float
        | Inch of Value: float
        | Millimeter of Value: float
        | Pica of Value: float

        static member TryFromJson(element: JsonElement) =
            let value =
                Json.tryGetDoubleProperty "value" element

            Json.tryGetStringProperty "unit" element
            |> Option.bind (function
                | "cm" -> value |> Option.map Unit.Centimeter
                | "pt" -> value |> Option.map Unit.Point
                | "in" -> value |> Option.map Unit.Inch
                | "mm" -> value |> Option.map Unit.Millimeter
                | "pc" -> value |> Option.map Unit.Pica
                | _ -> None)

        member u.ToDocObj() =
            match u with
            | Centimeter v -> MigraDocCore.DocumentObjectModel.Unit.FromCentimeter v
            | Point v -> MigraDocCore.DocumentObjectModel.Unit.FromPoint v
            | Inch v -> MigraDocCore.DocumentObjectModel.Unit.FromInch v
            | Millimeter v -> MigraDocCore.DocumentObjectModel.Unit.FromMillimeter v
            | Pica v -> MigraDocCore.DocumentObjectModel.Unit.FromPica v

    and [<RequireQualifiedAccess>] VerticalAlignment =
        | Top
        | Center
        | Bottom

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some VerticalAlignment.Top
            | 1 -> Some VerticalAlignment.Center
            | 2 -> Some VerticalAlignment.Bottom
            | _ -> None

        member va.ToDocObj() =
            match va with
            | VerticalAlignment.Top -> MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Top
            | VerticalAlignment.Center -> MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Center
            | VerticalAlignment.Bottom -> MigraDocCore.DocumentObjectModel.Tables.VerticalAlignment.Bottom

    and [<RequireQualifiedAccess>] ParagraphAlignment =
        | Left
        | Center
        | Right
        | Justify

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some Left
            | 1 -> Some Center
            | 2 -> Some Right
            | 3 -> Some Justify
            | _ -> None

        member pa.ToDocObj() =
            match pa with
            | Left -> MigraDocCore.DocumentObjectModel.ParagraphAlignment.Left
            | Center -> MigraDocCore.DocumentObjectModel.ParagraphAlignment.Center
            | Right -> MigraDocCore.DocumentObjectModel.ParagraphAlignment.Right
            | Justify -> MigraDocCore.DocumentObjectModel.ParagraphAlignment.Justify

    and [<RequireQualifiedAccess>] TableRowAlignment =
        | Left
        | Center
        | Right

        member tra.ToDocObj() =
            match tra with
            | Left -> MigraDocCore.DocumentObjectModel.Tables.RowAlignment.Left
            | Center -> MigraDocCore.DocumentObjectModel.Tables.RowAlignment.Center
            | Right -> MigraDocCore.DocumentObjectModel.Tables.RowAlignment.Right

    and Padding =
        { Top: Unit option
          Bottom: Unit option
          Left: Unit option
          Right: Unit option }

    and Margin =
        { Top: Unit option
          Bottom: Unit option
          Left: Unit option
          Right: Unit option }

    and [<RequireQualifiedAccess>] Color =
        | RGBA of R: byte * G: byte * B: byte * A: byte
        | Named of Name: string

        static member TryFromJson(element: JsonElement) =
            Json.tryGetStringProperty "type" element
            |> Option.bind (function
                | "rgba" ->
                    let value name =
                        Json.tryGetByteProperty name element
                        |> Option.defaultValue 0uy

                    Color.RGBA(value "r", value "g", value "b", value "a")
                    |> Some
                | "named" ->
                    Json.tryGetStringProperty "name" element
                    |> Option.map Color.Named
                | _ -> None)

        member c.ToDocObj() =
            match c with
            | RGBA (r, g, b, a) ->
                // Fully qualified parameters just to be safe.
                MigraDocCore.DocumentObjectModel.Color(a = a, r = r, g = g, b = b)
            | Named name -> MigraDocCore.DocumentObjectModel.Color.Parse name

    and Shading =
        { Color: Color }

        static member TryFromJson(element: JsonElement) =
            Json.tryGetProperty "color" element
            |> Option.bind Color.TryFromJson
            |> Option.map (fun c -> { Color = c })

        member s.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Shading()

            obj.Color <- s.Color.ToDocObj()

            obj

    and [<RequireQualifiedAccess>] BorderStyle =
        | None
        | Single
        | Dot
        | DashSmallGap
        | DashLargeGap
        | DashDot
        | DashDotDot

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some BorderStyle.None
            | 1 -> Some BorderStyle.Single
            | 2 -> Some BorderStyle.Dot
            | 3 -> Some BorderStyle.DashSmallGap
            | 4 -> Some BorderStyle.DashLargeGap
            | 5 -> Some BorderStyle.DashDot
            | 6 -> Some BorderStyle.DashDotDot
            | _ -> Option.None

        member bs.ToDocObj() =
            match bs with
            | BorderStyle.None -> MigraDocCore.DocumentObjectModel.BorderStyle.None
            | BorderStyle.Single -> MigraDocCore.DocumentObjectModel.BorderStyle.Single
            | BorderStyle.Dot -> MigraDocCore.DocumentObjectModel.BorderStyle.Dot
            | BorderStyle.DashSmallGap -> MigraDocCore.DocumentObjectModel.BorderStyle.DashSmallGap
            | BorderStyle.DashLargeGap -> MigraDocCore.DocumentObjectModel.BorderStyle.DashLargeGap
            | BorderStyle.DashDot -> MigraDocCore.DocumentObjectModel.BorderStyle.DashDot
            | BorderStyle.DashDotDot -> MigraDocCore.DocumentObjectModel.BorderStyle.DashDotDot

    and Borders =
        { Color: Color option
          Distance: Unit option
          Style: BorderStyle option
          Visible: bool option
          BordersCleared: bool option
          Width: Unit option
          DistanceFromTop: Unit option
          DistanceFromBottom: Unit option
          DistanceFromLeft: Unit option
          DistanceFromRight: Unit option
          Top: Border option
          Bottom: Border option
          Left: Border option
          Right: Border option
          DiagonalUp: Border option
          DiagonalDown: Border option }

        static member FromJson(element: JsonElement) =
            ({ Color =
                Json.tryGetProperty "color" element
                |> Option.bind Color.TryFromJson
               Distance =
                 Json.tryGetProperty "distance" element
                 |> Option.bind Unit.TryFromJson
               Style =
                 Json.tryGetIntProperty "style" element
                 |> Option.bind BorderStyle.Deserialize
               Visible = Json.tryGetBoolProperty "visible" element
               BordersCleared = Json.tryGetBoolProperty "bordersCleared" element
               Width =
                 Json.tryGetProperty "width" element
                 |> Option.bind Unit.TryFromJson
               DistanceFromTop =
                 Json.tryGetProperty "distanceFromTop" element
                 |> Option.bind Unit.TryFromJson
               DistanceFromBottom =
                 Json.tryGetProperty "distanceFromBottom" element
                 |> Option.bind Unit.TryFromJson
               DistanceFromLeft =
                 Json.tryGetProperty "distanceFromLeft" element
                 |> Option.bind Unit.TryFromJson
               DistanceFromRight =
                 Json.tryGetProperty "distanceFromRight" element
                 |> Option.bind Unit.TryFromJson
               Top =
                 Json.tryGetProperty "top" element
                 |> Option.map Border.FromJson
               Bottom =
                 Json.tryGetProperty "bottom" element
                 |> Option.map Border.FromJson
               Left =
                 Json.tryGetProperty "left" element
                 |> Option.map Border.FromJson
               Right =
                 Json.tryGetProperty "right" element
                 |> Option.map Border.FromJson
               DiagonalUp =
                 Json.tryGetProperty "diagonalUp" element
                 |> Option.map Border.FromJson
               DiagonalDown =
                 Json.tryGetProperty "diagonalDown" element
                 |> Option.map Border.FromJson }: Borders)

        member b.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Borders()

            match b.Color with
            | Some c -> obj.Color <- c.ToDocObj()
            | None -> ()

            match b.Distance with
            | Some d -> obj.Distance <- d.ToDocObj()
            | None -> ()

            match b.Style with
            | Some bs -> obj.Style <- bs.ToDocObj()
            | None -> ()

            match b.Visible with
            | Some v -> obj.Visible <- v
            | None -> ()

            match b.BordersCleared with
            | Some bc -> obj.BordersCleared <- bc
            | None -> ()

            match b.Width with
            | Some w -> obj.Width <- w.ToDocObj()
            | None -> ()

            match b.DistanceFromTop with
            | Some dft -> obj.DistanceFromTop <- dft.ToDocObj()
            | None -> ()

            match b.DistanceFromBottom with
            | Some dfb -> obj.DistanceFromBottom <- dfb.ToDocObj()
            | None -> ()

            match b.DistanceFromLeft with
            | Some dfl -> obj.DistanceFromLeft <- dfl.ToDocObj()
            | None -> ()

            match b.DistanceFromRight with
            | Some dfr -> obj.DistanceFromRight <- dfr.ToDocObj()
            | None -> ()

            match b.Top with
            | Some t -> obj.Top <- t.ToDocObj()
            | None -> ()

            match b.Bottom with
            | Some b -> obj.Bottom <- b.ToDocObj()
            | None -> ()

            match b.Left with
            | Some l -> obj.Left <- l.ToDocObj()
            | None -> ()

            match b.Right with
            | Some r -> obj.Right <- r.ToDocObj()
            | None -> ()

            match b.DiagonalUp with
            | Some du -> obj.DiagonalUp <- du.ToDocObj()
            | None -> ()

            match b.DiagonalDown with
            | Some dd -> obj.DiagonalDown <- dd.ToDocObj()
            | None -> ()

            obj

    and Border =
        { Color: Color option
          Visible: bool option
          Style: BorderStyle option
          Width: Unit option }

        static member FromJson(element: JsonElement) =
            ({ Color =
                Json.tryGetProperty "color" element
                |> Option.bind Color.TryFromJson
               Visible = Json.tryGetBoolProperty "visible" element
               Style =
                 Json.tryGetIntProperty "style" element
                 |> Option.bind BorderStyle.Deserialize
               Width =
                 Json.tryGetProperty "width" element
                 |> Option.bind Unit.TryFromJson }: Border)

        member b.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Border()

            match b.Color with
            | Some c -> obj.Color <- c.ToDocObj()
            | None -> ()

            match b.Visible with
            | Some v -> obj.Visible <- v
            | None -> ()

            match b.Style with
            | Some s -> obj.Style <- s.ToDocObj()
            | None -> ()

            match b.Width with
            | Some w -> obj.Width <- w.ToDocObj()
            | None -> ()

            obj

    and ParagraphFormat =
        { Alignment: ParagraphAlignment option
          Borders: Borders option
          Font: Font option
          Shading: Shading option
          KeepTogether: bool option
          LeftIndent: Unit option
          LineSpacing: Unit option
          ListInfo: ListInfo option
          OutlineLevel: OutlineLevel option
          RightIndent: Unit option
          SpaceBefore: Unit option
          SpaceAfter: Unit option
          // Tab stops
          WindowControl: bool option
          FirstLineIndent: Unit option
          KeepWithNext: bool option
          LineSpacingRule: LineSpacingRule option
          PageBreakBefore: bool option }

        static member Blank() =
            { Alignment = None
              Borders = None
              Font = None
              Shading = None
              KeepTogether = None
              LeftIndent = None
              LineSpacing = None
              ListInfo = None
              OutlineLevel = None
              RightIndent = None
              SpaceBefore = None
              SpaceAfter = None
              WindowControl = None
              FirstLineIndent = None
              KeepWithNext = None
              LineSpacingRule = None
              PageBreakBefore = None }

        static member FromJson(element: JsonElement) =
            ({ Alignment =
                Json.tryGetIntProperty "alignment" element
                |> Option.bind ParagraphAlignment.Deserialize
               Borders =
                 Json.tryGetProperty "borders" element
                 |> Option.map Borders.FromJson
               Font =
                 Json.tryGetProperty "font" element
                 |> Option.map Font.FromJson
               Shading =
                 Json.tryGetProperty "shading" element
                 |> Option.bind Shading.TryFromJson
               KeepTogether = Json.tryGetBoolProperty "keepTogether" element
               LeftIndent =
                 Json.tryGetProperty "leftIndent" element
                 |> Option.bind Unit.TryFromJson
               LineSpacing =
                 Json.tryGetProperty "lineSpacing" element
                 |> Option.bind Unit.TryFromJson
               ListInfo =
                 Json.tryGetProperty "listInfo" element
                 |> Option.map ListInfo.FromJson
               OutlineLevel = None
               RightIndent =
                 Json.tryGetProperty "rightIndent" element
                 |> Option.bind Unit.TryFromJson
               SpaceBefore =
                 Json.tryGetProperty "spaceBefore" element
                 |> Option.bind Unit.TryFromJson
               SpaceAfter =
                 Json.tryGetProperty "spaceAfter" element
                 |> Option.bind Unit.TryFromJson
               WindowControl = Json.tryGetBoolProperty "windowControl" element
               FirstLineIndent =
                 Json.tryGetProperty "firstLineIndent" element
                 |> Option.bind Unit.TryFromJson
               KeepWithNext = Json.tryGetBoolProperty "keepWithNext" element
               LineSpacingRule =
                 Json.tryGetIntProperty "lineSpacingRule" element
                 |> Option.bind LineSpacingRule.Deserialize
               PageBreakBefore = Json.tryGetBoolProperty "pageBreakBefore" element }: ParagraphFormat)

        member pf.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.ParagraphFormat()

            match pf.Alignment with
            | Some pa -> obj.Alignment <- pa.ToDocObj()
            | None -> ()

            match pf.Borders with
            | Some b -> obj.Borders <- b.ToDocObj()
            | None -> ()

            match pf.Font with
            | Some f -> obj.Font <- f.ToDocObj()
            | None -> ()

            match pf.Shading with
            | Some s -> obj.Shading <- s.ToDocObj()
            | None -> ()

            match pf.KeepTogether with
            | Some kt -> obj.KeepTogether <- kt
            | None -> ()

            match pf.LeftIndent with
            | Some li -> obj.LeftIndent <- li.ToDocObj()
            | None -> ()

            match pf.LineSpacing with
            | Some ls -> obj.LineSpacing <- ls.ToDocObj()
            | None -> ()

            match pf.ListInfo with
            | Some li -> obj.ListInfo <- li.ToDocObj()
            | None -> ()

            match pf.OutlineLevel with
            | Some ol -> obj.OutlineLevel <- ol.ToDocObj()
            | None -> ()

            match pf.RightIndent with
            | Some ri -> obj.RightIndent <- ri.ToDocObj()
            | None -> ()

            match pf.SpaceBefore with
            | Some sb -> obj.SpaceBefore <- sb.ToDocObj()
            | None -> ()

            match pf.SpaceAfter with
            | Some sa -> obj.SpaceAfter <- sa.ToDocObj()
            | None -> ()

            match pf.WindowControl with
            | Some wc -> obj.WidowControl <- wc
            | None -> ()

            match pf.FirstLineIndent with
            | Some fli -> obj.FirstLineIndent <- fli.ToDocObj()
            | None -> ()

            match pf.KeepWithNext with
            | Some kwn -> obj.KeepWithNext <- kwn
            | None -> ()

            match pf.LineSpacingRule with
            | Some ls -> obj.LineSpacingRule <- ls.ToDocObj()
            | None -> ()

            match pf.PageBreakBefore with
            | Some pbb -> obj.PageBreakBefore <- pbb
            | None -> ()

            obj

    and Font =
        { Bold: bool option
          Color: Color option
          Italic: bool option
          Name: string option
          Size: Unit option
          Subscript: bool option
          Superscript: bool option
          Underline: Underline option }

        static member Default() =
            { Bold = None
              Color = None
              Italic = None
              Name = None
              Size = None
              Subscript = None
              Superscript = None
              Underline = None }

        static member FromJson(element: JsonElement) =
            ({ Bold = Json.tryGetBoolProperty "bold" element
               Color =
                 Json.tryGetProperty "color" element
                 |> Option.bind Color.TryFromJson
               Italic = Json.tryGetBoolProperty "italic" element
               Name = Json.tryGetStringProperty "name" element
               Size =
                 Json.tryGetProperty "size" element
                 |> Option.bind Unit.TryFromJson
               Subscript = Json.tryGetBoolProperty "subscript" element
               Superscript = Json.tryGetBoolProperty "superscript" element
               Underline =
                 Json.tryGetIntProperty "underline" element
                 |> Option.bind Underline.Deserialize }: Font)

        member f.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.Font()

            match f.Bold with
            | Some v -> obj.Bold <- v
            | None -> ()

            match f.Color with
            | Some c -> obj.Color <- c.ToDocObj()
            | None -> ()

            match f.Italic with
            | Some i -> obj.Italic <- i
            | None -> ()

            match f.Name with
            | Some n -> obj.Name <- n
            | None -> ()

            match f.Size with
            | Some s -> obj.Size <- s.ToDocObj()
            | None -> ()

            match f.Subscript with
            | Some s -> obj.Subscript <- s
            | None -> ()

            match f.Superscript with
            | Some s -> obj.Superscript <- s
            | None -> ()

            match f.Underline with
            | Some u -> obj.Underline <- u.ToDocObj()
            | None -> ()

            obj

    and [<RequireQualifiedAccess>] Underline =
        | None
        | Single
        | Words
        | Dotted
        | Dash
        | DotDash
        | DotDotDash

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some Underline.None
            | 1 -> Some Underline.Single
            | 2 -> Some Underline.Words
            | 3 -> Some Underline.Dotted
            | 4 -> Some Underline.Dash
            | 5 -> Some Underline.DotDash
            | 6 -> Some Underline.DotDotDash
            | _ -> Option.None

        member ul.ToDocObj() =
            match ul with
            | None -> MigraDocCore.DocumentObjectModel.Underline.None
            | Single -> MigraDocCore.DocumentObjectModel.Underline.Single
            | Words -> MigraDocCore.DocumentObjectModel.Underline.Words
            | Dotted -> MigraDocCore.DocumentObjectModel.Underline.Dotted
            | Dash -> MigraDocCore.DocumentObjectModel.Underline.Dash
            | DotDash -> MigraDocCore.DocumentObjectModel.Underline.DotDash
            | DotDotDash -> MigraDocCore.DocumentObjectModel.Underline.DotDotDash


    and OutlineLevel =
        | BodyText
        | Level1
        | Level2
        | Level3
        | Level4
        | Level5
        | Level6
        | Level7
        | Level8
        | Level9

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some OutlineLevel.BodyText
            | 1 -> Some OutlineLevel.Level1
            | 2 -> Some OutlineLevel.Level2
            | 3 -> Some OutlineLevel.Level3
            | 4 -> Some OutlineLevel.Level4
            | 5 -> Some OutlineLevel.Level5
            | 6 -> Some OutlineLevel.Level6
            | 7 -> Some OutlineLevel.Level7
            | 8 -> Some OutlineLevel.Level8
            | 9 -> Some OutlineLevel.Level9
            | _ -> None

        member ol.ToDocObj() =
            match ol with
            | BodyText -> MigraDocCore.DocumentObjectModel.OutlineLevel.BodyText
            | Level1 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level1
            | Level2 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level2
            | Level3 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level3
            | Level4 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level4
            | Level5 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level5
            | Level6 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level6
            | Level7 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level7
            | Level8 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level8
            | Level9 -> MigraDocCore.DocumentObjectModel.OutlineLevel.Level9

    and ListInfo =
        { ListType: ListType option
          NumberPosition: Unit option
          ContinuePreviousList: bool option }

        static member FromJson(element: JsonElement) =
            ({ ListType =
                Json.tryGetIntProperty "listType" element
                |> Option.bind ListType.Deserialize
               NumberPosition =
                 Json.tryGetProperty "numberPosition" element
                 |> Option.bind Unit.TryFromJson
               ContinuePreviousList = Json.tryGetBoolProperty "continuePreviousList" element }: ListInfo)

        member li.ToDocObj() =
            let obj =
                MigraDocCore.DocumentObjectModel.ListInfo()

            match li.ListType with
            | Some lt -> obj.ListType <- lt.ToDocObj()
            | None -> ()

            match li.NumberPosition with
            | Some np -> obj.NumberPosition <- np.ToDocObj()
            | None -> ()

            match li.ContinuePreviousList with
            | Some cpl -> obj.ContinuePreviousList <- cpl
            | None -> ()

            obj

    and ListType =
        | BulletList1
        | BulletList2
        | BulletList3
        | NumberList1
        | NumberList2
        | NumberList3

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some BulletList1
            | 1 -> Some BulletList2
            | 2 -> Some BulletList3
            | 3 -> Some NumberList1
            | 4 -> Some NumberList2
            | 5 -> Some NumberList3
            | _ -> None

        member lt.ToDocObj() =
            match lt with
            | BulletList1 -> MigraDocCore.DocumentObjectModel.ListType.BulletList1
            | BulletList2 -> MigraDocCore.DocumentObjectModel.ListType.BulletList2
            | BulletList3 -> MigraDocCore.DocumentObjectModel.ListType.BulletList3
            | NumberList1 -> MigraDocCore.DocumentObjectModel.ListType.BulletList1
            | NumberList2 -> MigraDocCore.DocumentObjectModel.ListType.BulletList2
            | NumberList3 -> MigraDocCore.DocumentObjectModel.ListType.BulletList3

    and [<RequireQualifiedAccess>] LineSpacingRule =
        | Single
        | OnePtFive
        | Double
        | AtLeast
        | Exactly
        | Multiple

        static member Deserialize(value: int) =
            match value with
            | 0 -> Some LineSpacingRule.Single
            | 1 -> Some LineSpacingRule.OnePtFive
            | 2 -> Some LineSpacingRule.Double
            | 3 -> Some LineSpacingRule.AtLeast
            | 4 -> Some LineSpacingRule.Exactly
            | 5 -> Some LineSpacingRule.Multiple
            | _ -> None

        member s.ToDocObj() =
            match s with
            | Single -> MigraDocCore.DocumentObjectModel.LineSpacingRule.Single
            | OnePtFive -> MigraDocCore.DocumentObjectModel.LineSpacingRule.OnePtFive
            | Double -> MigraDocCore.DocumentObjectModel.LineSpacingRule.Double
            | AtLeast -> MigraDocCore.DocumentObjectModel.LineSpacingRule.AtLeast
            | Exactly -> MigraDocCore.DocumentObjectModel.LineSpacingRule.Exactly
            | Multiple -> MigraDocCore.DocumentObjectModel.LineSpacingRule.Multiple