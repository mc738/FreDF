namespace FreDF.Core

open System.IO

module PdfRenderer =
    
    let init (style: PDFStyle) =
        let document = MigraDocCore.DocumentObjectModel.Document()
        style.DefineDocumentStyle document
        document
        
    
    let build (styles: PDFStyle) (document: Structure.PdfDocument) =
        let pdf = init styles
        
        document.Sections
        |> List.iter (fun s -> pdf.Add(s.ToDocObj()))
        
        pdf
    
    let run (styles: PDFStyle) (savePath: string) (document: Structure.PdfDocument) =
        let pdf =  build styles document
        
        let renderer = MigraDocCore.Rendering.PdfDocumentRenderer(true)
        renderer.Document <- pdf
        renderer.RenderDocument()
        renderer.PdfDocument.Save(savePath)
        
    let copyToStream (styles: PDFStyle) (stream: Stream) (document: Structure.PdfDocument) =
        let pdf =  build styles document
        
        let renderer = MigraDocCore.Rendering.PdfDocumentRenderer(true)
        renderer.Document <- pdf
        renderer.RenderDocument()
        renderer.PdfDocument.Save(stream)
        
        
    
        
        

