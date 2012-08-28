<%@page import="java.net.*,java.io.*,com.itextpdf.text.Chunk,com.itextpdf.text.Font.*,com.itextpdf.text.pdf.*,com.itextpdf.text.*;"%>
<%
/****
*Author: 	Doug Weems
*license:	opensource
*date: 	November 02, 2011
*****/
%>
<%
    //Defaults, override these in query string or form
    //There is a size limit in query string, 
    //ie. http://tms.bennettware.com/TMS/Application/Help/quickPDF.jsp?doctype=makepdf&header=cat,dog,fish,plane,my train,snake&details=car,1,2,3,4,5
    //but not in form POST.
    String header = "One, Two";
    String details = "1,2\n3,4\n";
    
    //optional
    String sTitle = "";
    Float headerfill = 0.85f;  //This is fill color for header column
    int headerFontSize = 12;
    int detailFontSize = 8;
    String sPage = "A4";
    String doctype = "";
    
    //Page Options
    com.itextpdf.text.Rectangle pageSize = PageSize.A4;  
    if (request.getParameter("sPage") != null)
         sPage = request.getParameter("sPage");
    if(sPage.equals("A4.rotate"))
        pageSize = PageSize.A4.rotate();
    else if(sPage.equals("LETTER"))
        pageSize = PageSize.LETTER;
    else if(sPage.equals("LEGAL"))
        pageSize = PageSize.LEGAL;        
    else if(sPage.equals("LEGAL.rotate"))
        pageSize = PageSize.LEGAL.rotate();    
    
    //PDF or Exel?
    if (request.getParameter("doctype") != null)
         doctype = request.getParameter("doctype");

    if(doctype.equals("makepdf")){
        response.setContentType("application/pdf");
        Document document = new Document(pageSize);
    
        //Read in Title
        if (request.getParameter("Title") != null)
             sTitle = request.getParameter("Title");
        //Read in Header Font Size
        if (request.getParameter("headerFontSize") != null)
             headerFontSize = Integer.parseInt(request.getParameter("headerFontSize"));
        //Read in Details Font Size
        if (request.getParameter("detailFontSize") != null)
             detailFontSize = Integer.parseInt(request.getParameter("detailFontSize"));
        
        Font arial = FontFactory.getFont("Arial", headerFontSize);
        arial.setStyle(Font.BOLD);
        Font arialdetail = FontFactory.getFont("Arial", detailFontSize);
        try{
                ByteArrayOutputStream buffer = new ByteArrayOutputStream();
                PdfWriter.getInstance(document, buffer);
                document.open();
    
                //Add Title
                Font font = new Font(FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.DARK_GRAY);
                Chunk id = new Chunk(sTitle, font);
                id.setTextRise(6);
                document.add(id);
                    
                 //Read in Header
                 if (request.getParameter("header") != null)
                     header = URLDecoder.decode(request.getParameter("header"), "UTF-8");
    
                 //Number of Columns
                 String[] arr = header.split(",");
                 int columns = arr.length;
                 
                 //Read in Details
                 if (request.getParameter("details") != null)
                     details = URLDecoder.decode(request.getParameter("details"), "UTF-8");
                  //Read in Header fill
                 if (request.getParameter("headerfill") != null){
                     String floating = request.getParameter("headerfill");
                     headerfill = new Float(floating);
                 }
                 //Create Table 
                 PdfPTable table = new PdfPTable(columns);
                 table.setWidthPercentage(100);
                
                 PdfPCell cell;
                
                 //Header Row
                 String[] st = header.split(",");
                 for(int i=0; i < st.length; i++){
                    cell = new PdfPCell(new Phrase(st[i], arial));
                    cell.setGrayFill(headerfill);
                    cell.setNoWrap(true);
                    table.addCell(cell);
                 }            
                
                //Detail Rows
                String[] rows = details.split("\n");
                String[] innercolumns;
                for(int i=0; i < rows.length; i++){
                     innercolumns = rows[i].split(",");
                    for(int y=0; y < innercolumns.length; y++){
                         table.addCell(new Phrase(innercolumns[y], arialdetail));
                    }  
                }              
                    
                // Code 5
                document.add(table);		
                document.close();
    
                DataOutput dataOutput = new DataOutputStream(response.getOutputStream());
                byte[] bytes = buffer.toByteArray();
                response.setContentLength(bytes.length);
                for(int i = 0; i < bytes.length; i++){
                    dataOutput.writeByte(bytes[i]);
                }
                    
            }catch(DocumentException e){
                e.printStackTrace();
            }
    }
    else {
        response.reset();
        response.setHeader("Content-type","application/xls");
        response.setHeader("Content-disposition","inline; filename=excel.csv");
         //Read in Header
         if (request.getParameter("header") != null)
             header = URLDecoder.decode(request.getParameter("header"), "UTF-8").trim();
         out.println(header);
         //Read in Details
         if (request.getParameter("details") != null)
             details = URLDecoder.decode(request.getParameter("details"), "UTF-8").trim();
         out.println(details);
    }
%>