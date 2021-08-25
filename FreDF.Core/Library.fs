namespace FreDF.Core

open System.Text.Json
open System.Text.Json.Serialization
open MigraDocCore.DocumentObjectModel
open MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes
open MigraDocCore.DocumentObjectModel.Shapes
open MigraDocCore.Rendering
open MigraDocCore.Rendering
open PdfSharpCore.Pdf

type PDFColor =
    { [<JsonPropertyName("r")>]
      R: byte
      [<JsonPropertyName("g")>]
      G: byte
      [<JsonPropertyName("b")>]
      B: byte
      [<JsonPropertyName("a")>]
      A: byte }

type PDFFontStyle =
    { [<JsonPropertyName("name")>]
      Name: string
      [<JsonPropertyName("bold")>]
      Bold: bool
      [<JsonPropertyName("italic")>]
      Italic: bool
      [<JsonPropertyName("size")>]
      Size: string
      [<JsonPropertyName("subscript")>]
      Subscript: bool
      [<JsonPropertyName("superscript")>]
      SuperScript: bool
      [<JsonPropertyName("underline")>]
      Underline: bool
      [<JsonPropertyName("color")>]
      Color: PDFColor }

type PDFParagraphStyle =
    { [<JsonPropertyName("firstLineIndent")>]
      FirstLineIndent: string
      [<JsonPropertyName("keepTogether")>]
      KeepTogether: bool
      [<JsonPropertyName("keepWithNext")>]
      KeepWithNext: bool
      [<JsonPropertyName("leftIndent")>]
      LeftIndent: string
      [<JsonPropertyName("lineSpacing")>]
      LineSpacing: string
      [<JsonPropertyName("pageBreakBefore")>]
      PageBreakBefore: bool
      [<JsonPropertyName("rightIndent")>]
      RightIndent: string
      [<JsonPropertyName("spaceAfter")>]
      SpaceAfter: string
      [<JsonPropertyName("spaceBefore")>]
      SpaceBefore: string
      [<JsonPropertyName("align")>]
      Align: int }

type PDFHeaderStyle =
    { [<JsonPropertyName("name")>]
      Name: string
      [<JsonPropertyName("paragraph")>]
      Paragraph: PDFParagraphStyle
      [<JsonPropertyName("font")>]
      Font: PDFFontStyle }

    member s.DefineStyle(document: Document, name: string) =
        let style = document.AddStyle(name, "Normal")

        let toUnit (v: string) = Unit.op_Implicit v
        
        style.Font.Bold <- s.Font.Bold
        style.Font.Color <- Color(s.Font.Color.A, s.Font.Color.R, s.Font.Color.G, s.Font.Color.B)
        style.Font.Italic <- s.Font.Italic
        style.Font.Name <- s.Font.Name
        style.Font.Size <- s.Font.Size |> toUnit
        style.Font.Subscript <- s.Font.Subscript
        style.Font.Superscript <- s.Font.SuperScript
        // style.Font.Underline <- s.Font.Underline
        style.ParagraphFormat.FirstLineIndent <- s.Paragraph.FirstLineIndent |> toUnit 
        style.ParagraphFormat.KeepTogether <- s.Paragraph.KeepTogether
        style.ParagraphFormat.KeepWithNext <- s.Paragraph.KeepWithNext
        style.ParagraphFormat.LeftIndent <- s.Paragraph.LeftIndent |> toUnit
        style.ParagraphFormat.LineSpacing <- s.Paragraph.LeftIndent |> toUnit
        style.ParagraphFormat.PageBreakBefore <- s.Paragraph.PageBreakBefore
        style.ParagraphFormat.RightIndent <- s.Paragraph.RightIndent |> toUnit
        style.ParagraphFormat.SpaceAfter <- s.Paragraph.SpaceAfter |> toUnit
        style.ParagraphFormat.SpaceBefore <- s.Paragraph.SpaceBefore |> toUnit
        style.ParagraphFormat.Alignment <- enum<ParagraphAlignment> s.Paragraph.Align


type PDFStyle =
    { [<JsonPropertyName("name")>]
      Name: string
      [<JsonPropertyName("paragraph")>]
      Paragraph: PDFParagraphStyle
      [<JsonPropertyName("font")>]
      Font: PDFFontStyle
      [<JsonPropertyName("h1")>]
      H1: PDFHeaderStyle
      [<JsonPropertyName("h2")>]
      H2: PDFHeaderStyle
      [<JsonPropertyName("h3")>]
      H3: PDFHeaderStyle
      [<JsonPropertyName("h4")>]
      H4: PDFHeaderStyle
      [<JsonPropertyName("h5")>]
      H5: PDFHeaderStyle
      [<JsonPropertyName("h6")>]
      H6: PDFHeaderStyle }

    member s.DefineStyle(document: Document, defaultStyle: bool, name: string) =
        let style =
            match defaultStyle with
            | true -> document.Styles.["Normal"]
            | false -> document.AddStyle(name, "Normal")

        let toUnit (v: string) = Unit.op_Implicit v
        
        style.Font.Bold <- s.Font.Bold
        style.Font.Color <- Color(s.Font.Color.A, s.Font.Color.R, s.Font.Color.G, s.Font.Color.B)
        style.Font.Italic <- s.Font.Italic
        style.Font.Name <- s.Font.Name
        style.Font.Size <- s.Font.Size |> toUnit
        style.Font.Subscript <- s.Font.Subscript
        style.Font.Superscript <- s.Font.SuperScript
        // style.Font.Underline <- s.Font.Underline
        style.ParagraphFormat.FirstLineIndent <- s.Paragraph.FirstLineIndent |> toUnit 
        style.ParagraphFormat.KeepTogether <- s.Paragraph.KeepTogether
        style.ParagraphFormat.KeepWithNext <- s.Paragraph.KeepWithNext
        style.ParagraphFormat.LeftIndent <- s.Paragraph.LeftIndent |> toUnit
        style.ParagraphFormat.LineSpacing <- s.Paragraph.LeftIndent |> toUnit
        style.ParagraphFormat.PageBreakBefore <- s.Paragraph.PageBreakBefore
        style.ParagraphFormat.RightIndent <- s.Paragraph.RightIndent |> toUnit
        style.ParagraphFormat.SpaceAfter <- s.Paragraph.SpaceAfter |> toUnit
        style.ParagraphFormat.SpaceBefore <- s.Paragraph.SpaceBefore |> toUnit
        style.ParagraphFormat.Alignment <- enum<ParagraphAlignment> s.Paragraph.Align

    member s.DefineDocumentStyle(document: Document) =
        s.DefineStyle(document, true, "")
        s.H1.DefineStyle(document, "Heading1")
        s.H2.DefineStyle(document, "Heading2")
        //s.H3.DefineStyle(document, "Heading3")
        //s.H4.DefineStyle(document, "Heading4")
        //s.H5.DefineStyle(document, "Heading5")
        //s.H6.DefineStyle(document, "Heading6")
    

type PDF(document: Document) =
        
    let mutable currentSection: Section = document.AddSection()

    member _.AddSection(isPortrait: bool) =
        currentSection <- document.AddSection()
        currentSection.PageSetup.Orientation <- if isPortrait then Orientation.Portrait else Orientation.Landscape

    member pdf.AddHeader1(text: string, newSection: bool) =
        if newSection then currentSection <- document.AddSection()
        let paragraph = currentSection.AddParagraph()
        paragraph.AddText(text) |> ignore
        paragraph.Style = "Heading1" |> ignore
        
    member pdf.AddHeader2(text: string, newSection: bool) =
        if newSection then currentSection <- document.AddSection()
        let paragraph = currentSection.AddParagraph()
        paragraph.AddText(text) |> ignore
        paragraph.Style = "Heading2" |> ignore
        
    member pdf.AddParagraph(text: string, newSection: bool) =
        if newSection then currentSection <- document.AddSection()
        let paragraph = currentSection.AddParagraph()
        paragraph.AddText(text) |> ignore
        paragraph.Style = "Normal" |> ignore
    
    member pdf.AddStyledParagraph(text: string, style: string, newSection: bool) =
        if newSection then currentSection <- document.AddSection()
        let paragraph = currentSection.AddParagraph()
        paragraph.AddText(text) |> ignore
        paragraph.Style = style |> ignore
        
    member _.AddImage(path: string, width: string, newSection: bool, wrapInParagraph: bool) =
        if newSection then currentSection <- document.AddSection()
        
        let image =
            match wrapInParagraph with
            | true ->
                let paragraph = currentSection.AddParagraph()
                paragraph.AddImage(ImageSource.FromFile(path))
            | false -> currentSection.AddImage(ImageSource.FromFile(path))
            
        image.Width = Unit.op_Implicit width |> ignore     
        
    member _.Render(location: string) =
        let renderer = PdfDocumentRenderer(true)
        renderer.Document <- document
        renderer.RenderDocument()
        renderer.PdfDocument.Save(location)
    
