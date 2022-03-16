
// **************************************************************************
//  SQL2XLSX - Create Excel XLSX from SQL
//  copyright('Giuseppe Costagliola - giuseppe.costagliola@gmail.com - 2017')
//  extended Erik Mennen 2022 
//  -------------------------------------------------------------------------

//  APACHE POI 3.9:
//   - poi-3.9.jar
//   - poi-ooxml-3.9.jar
//   - poi-ooxml-schemas-3.9.jar
//
//   DOM4J 1.6.1 **:
//   - dom4j-1.6.1.jar
//
//   XMLBEANS 2.5.0 **:
//   - jsr173_1.0_api.jar
//   - xbean.jar
//
//   ** = extra JAR files required for XSSF (Office XML) format.
//
// ********************************************************************
//  "This product includes software developed by the
//   Apache Software Foundation (http://www.apache.org/)."
// --------------------------------------------------------------------
//  THIS UTILITY IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED
//  WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
//  OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
//  DISCLAIMED.  IN NO EVENT SHALL THE AUTHOR OF THIS UTILITY OR
//  ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
//  SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
//  LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
//  USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
//  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
//  OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT
//  OF THE USE OF THIS UTILITY, EVEN IF ADVISED OF THE POSSIBILITY OF
//  SUCH DAMAGE.
// ********************************************************************


import java.sql.*;
import java.io.*;
import java.util.Date;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Font;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.CellType;

public class SQL2XLXNS {

  static class ColHeader
  {
    String label;
    int width;
    boolean wrap;
  }

  public static void main (String[] parameters) {

    if (parameters.length == 23) {

        String SQLID       = parameters[0];
        String TOXLS       = parameters[1];
        String FROMXLS     = parameters[2];
        String SHEETNAME   = parameters[3];
        String COLHDRS     = parameters[4];
        String TITLE       = parameters[5];
        String TITLECOLS   = parameters[6];
        String TITLEALIGN  = parameters[7];
        String LTRIM       = parameters[8];
        String WRTZEROC    = parameters[9];
        String FREEZE      = parameters[10];
        String FITPRT      = parameters[11];
        String GRID        = parameters[12];
        String ORIENT      = parameters[13];
        String BLANKROWS   = parameters[14];
        String NAMING      = parameters[15];
        String JLIBL       = parameters[16];
        String TRANBIN     = parameters[17];
        String DRIVER      = parameters[18];
        String SYSTEM      = parameters[19];
        String USER        = parameters[20];
        String PASSWORD    = parameters[21];
        String DEBUG       = parameters[22];

        Date dateStarted = new Date();
        long timeStarted = dateStarted.getTime();

      try {

        // check parameters
        if (DRIVER.equals("NATIVE") & !SYSTEM.equals("LOCALHOST")) {
          System.out.println("ERROR: Driver/System inconsistency");
           System.exit(1);
        }

        // get the SQL statement
        FileInputStream fin = new FileInputStream(SQLID);
        ByteArrayOutputStream bout = new ByteArrayOutputStream();
        copySQL(fin,bout);
        fin.close();
        String sqlStatement = bout.toString();
        if (DEBUG.equals("Y")) System.out.println(sqlStatement);

        // Create a properties object and set the properties for the connection.
        Properties p = new Properties();

        if ((DRIVER.equals("NATIVE")) | (DRIVER.equals("JT400"))) {
          if (COLHDRS.equals("L"))
             p.put("extended metadata", "true");
          else
             p.put("extended metadata", "false");
          if (NAMING.equals("SYS"))
             p.put("naming", "system");
          else
             p.put("naming", "sql");
          if (!JLIBL.equals("DFT"))
            p.put("libraries", JLIBL);
          if (TRANBIN.equals("Y"))
            p.put("translate binary", "true");
        }

        if (!USER.equals(" ")) p.put("user", USER.trim());
        if (!PASSWORD.equals(" ")) p.put("password", PASSWORD.trim());

        // initialize connection object
        Connection connection = null;

        try {

            String jdbcDriver = null;
            String jdbcConnection = null;
            if (DRIVER.equals("NATIVE")) {
              jdbcDriver = "com.ibm.db2.jdbc.app.DB2Driver";
              jdbcConnection = "jdbc:db2://localhost";
            } else if (DRIVER.equals("JT400")) {
              jdbcDriver = "com.ibm.as400.access.AS400JDBCDriver";
              jdbcConnection = "jdbc:as400://" + SYSTEM;
            } else if (DRIVER.equals("JTDS")) {
              jdbcDriver = "net.sourceforge.jtds.jdbc.Driver";
              jdbcConnection = "jdbc:jtds:sqlserver://" + SYSTEM;
            }
            if (DEBUG.equals("Y")) {
              System.out.println("Connecting to " + SYSTEM.trim() +
                                 " with \"" + jdbcDriver + "\"");
              if (!USER.equals(" "))
                System.out.println("User:" + USER.trim() +
                                   "\nPassword: " + PASSWORD);
            }

            // Load Java JDBC driver and connect
            Class.forName(jdbcDriver);
            connection = DriverManager.getConnection (jdbcConnection, p);

            // get information about the connection
            DatabaseMetaData dmd = connection.getMetaData ();

            // execute the query.
            Statement select = connection.createStatement ();
            ResultSet rs = select.executeQuery (sqlStatement);

            // sheet row and column number
            int rowNum = 0;
            short colNum = 0;
            String nullValue = "<null>";
            String emptyValue = "";

            // create workbook and sheet
            XSSFWorkbook wb = null;
            XSSFSheet sheet = null;

            // create a new workbook
            if (FROMXLS.equals(" ")) {
             wb = new XSSFWorkbook();
             int iBlankRows = Integer.valueOf(BLANKROWS).intValue();
             if (SHEETNAME.equals(" ")) sheet = wb.createSheet();
             else sheet = wb.createSheet(SHEETNAME);
            // open an existing workbook
            } else {
			     wb = new XSSFWorkbook(FROMXLS);
              int iBlankRows = Integer.valueOf(BLANKROWS).intValue();
              // get sheet
              if (SHEETNAME.equals(" ")) {
                 sheet = wb.getSheetAt(0);
                 // clear sheet
                 if (iBlankRows >= 0) {
                    rowNum = sheet.getLastRowNum() + iBlankRows + 1;
                 }
                 else {
                    wb.removeSheetAt(0);
                    sheet = wb.createSheet();
                    rowNum = 0;
                 }
              }
              else {
                 sheet = wb.getSheet(SHEETNAME);
              }
              // add new sheet
              if (sheet == null) {
                 sheet = wb.createSheet(SHEETNAME);
                 rowNum = 0;
              } else {
                // clear sheet
                if (iBlankRows >= 0) {
                  rowNum = sheet.getLastRowNum() + iBlankRows + 1;
                }
                else {
                  int sheetIndex = wb.getSheetIndex(SHEETNAME);
                  if (sheetIndex >= 0) {
                    wb.removeSheetAt(sheetIndex);
                    sheet = wb.createSheet(SHEETNAME);
                    rowNum = 0;
                  }
                }
              }
            }

            // declare row, cell reference
            XSSFRow row = null;
            XSSFCell cell = null;

            // create Font object e set it Bold
            XSSFFont font = wb.createFont();
            font.setBold(true);

            // create some text styles
            XSSFCellStyle styleBold = wb.createCellStyle();
            styleBold.setFont(font);
            XSSFCellStyle styleAlignC = wb.createCellStyle();
            styleAlignC.setAlignment(HorizontalAlignment.CENTER);
            styleAlignC.setFont(font);

            // create *d Style
            XSSFCellStyle style0d = wb.createCellStyle();
            XSSFCellStyle style1d = wb.createCellStyle();
            XSSFCellStyle style2d = wb.createCellStyle();
            XSSFCellStyle style3d = wb.createCellStyle();
            XSSFCellStyle style4d = wb.createCellStyle();

            // create a DataFormatter, a User Format and set the Style
            XSSFDataFormat df = wb.createDataFormat();
            style0d.setDataFormat(df.getFormat("#,##0"));
            style1d.setDataFormat(df.getFormat("#,##0.0"));
            style2d.setDataFormat(df.getFormat("#,##0.00"));
            style3d.setDataFormat(df.getFormat("#,##0.000"));
            style4d.setDataFormat(df.getFormat("#,##0.0000"));

            // create a Date-Time style
            XSSFCellStyle styleDate = wb.createCellStyle();
            styleDate.setDataFormat(df.getFormat("m/d/yy"));
            XSSFCellStyle styleTime = wb.createCellStyle();
            styleTime.setDataFormat(df.getFormat("h:mm:ss"));
            styleTime.setAlignment(HorizontalAlignment.CENTER);

            // get information about the result set.
            ResultSetMetaData rsmd = rs.getMetaData ();
            int columnCount = rsmd.getColumnCount ();
            int[] columnWitdh = new int[columnCount];

            // If requested, write a new row for sheet title
            if (! TITLE.equals(" ")) {
               // create a style for sheet header
               XSSFFont fontTitle = wb.createFont();
               fontTitle.setFontHeightInPoints((short) 12);
               fontTitle.setBold(true);
               XSSFCellStyle styleTitle = wb.createCellStyle();
               styleTitle.setFont(fontTitle);
               styleTitle.setWrapText(false);
               // styleTitle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
               styleTitle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
               XSSFCellStyle styleTitleC = wb.createCellStyle();
               styleTitleC.setFont(fontTitle);
               styleTitleC.setAlignment(HorizontalAlignment.CENTER);
               styleTitleC.setWrapText(false);
               // styleTitleC.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
               styleTitleC.setFillPattern(FillPatternType.SOLID_FOREGROUND);
               // create a header row
               row = sheet.createRow(rowNum++);
               // add first column
               cell = row.createCell((short) 0);
               cell.setCellType(CellType.STRING);
               cell.setCellValue(TITLE);
               // set header alignment
               if (TITLEALIGN.equals("N"))
                 cell.setCellStyle(styleTitle);
               else
                 cell.setCellStyle(styleTitleC);
               // set number of columns
               int iTitleCols = Integer.valueOf(TITLECOLS).intValue();
               if (iTitleCols == -1) iTitleCols = columnCount;
               if (iTitleCols > 1) {  // add more columns
                 for (colNum = 1; colNum < iTitleCols; colNum++) {
                   cell = row.createCell(colNum);
                 }
                 sheet.addMergedRegion(new CellRangeAddress(rowNum-1, (short) 0,
                                       rowNum-1, (short) iTitleCols));
               }
            }

            // write the column headers row
            if (!COLHDRS.equals("N")) {
              row = sheet.createRow(rowNum++);
            }
            // set the column length and adjust
            for (colNum = (short) 0; colNum < columnCount; colNum++) {
                columnWitdh[colNum] = rsmd.getPrecision (colNum+1);
                if ((rsmd.getColumnType (colNum+1) == Types.DOUBLE) |
                    (rsmd.getColumnType (colNum+1) == Types.BIGINT))
                  columnWitdh[colNum] = 10;
                else if (rsmd.getColumnType (colNum+1) == Types.CHAR)
                  columnWitdh[colNum] += 2;

                // write column header
                if (!COLHDRS.equals("N")) {
                  ColHeader hdr = null;
                  String colHdg = null;
                  cell = row.createCell(colNum);  // create a cell
                  // label/colhdg
                  if (COLHDRS.equals("L")) {
                    colHdg =  rsmd.getColumnLabel(colNum+1);
                    int colHdgL = colHdg.trim().length();
                    if (colHdgL > 40)
                        colHdg = colHdg.substring(0,20).trim() + "\\" +
                                 colHdg.substring(20,40).trim() +  "\\" +
                                 colHdg.substring(40);
                    else if (colHdgL > 20)
                        colHdg = colHdg.substring(0,20).trim() + "\\" +
                                 colHdg.substring(20);
                    hdr = setColHeader(colNum, colHdg.trim(), "\\", "\012");
                  }
                  // field name
                  else {
                    colHdg =  rsmd.getColumnLabel(colNum+1);
                    if (colHdg.substring(0,1).equals("\""))
                       colHdg = colHdg.substring(1,colHdg.trim().length()-2).trim();
                    hdr = setColHeader(colNum, colHdg, "\\", "\012");
                    hdr.width += 1;
                  }
                  // take le longest and set the wrap attribute
                  if (hdr.width > columnWitdh[colNum]) columnWitdh[colNum] = hdr.width;
                  if (hdr.wrap) styleBold.setWrapText(true);
                  if (hdr.wrap) styleAlignC.setWrapText(true);
                  // write cell
                  cell.setCellType(CellType.STRING);
                  cell.setCellValue(hdr.label);
                  int colType = rsmd.getColumnType (colNum+1);
                  if (colType == 3)   // if numeric align center
                    cell.setCellStyle(styleAlignC);
                  else
                    cell.setCellStyle(styleBold);
              }
            }

            // freeze headers and repeat rows
            if (!COLHDRS.equals("N")) {
               if ((FROMXLS.equals(" ")) & (FREEZE.equals("Y")))
                 sheet.createFreezePane( 0, rowNum, 0, rowNum );
                 int sheetIndex;
                 if (!SHEETNAME.equals(" ")) sheetIndex = wb.getSheetIndex(SHEETNAME);
                 else sheetIndex = wb.getNumberOfSheets() - 1;
                //wb.setRepeatingRowsAndColumns (sheetIndex, -1, -1, 0, rowNum - 1);
            }

            // fit sheet
            if (FITPRT.equals("Y")) {
              XSSFPrintSetup ps = sheet.getPrintSetup();
              sheet.setAutobreaks(true);
              ps.setFitHeight((short)1);
              ps.setFitWidth((short)1);
            }

            // orientation
            if (ORIENT.equals("L"))
              sheet.getPrintSetup().setLandscape(true);

            // fetch records
            while (rs.next ()) {
              row = sheet.createRow(rowNum++);    // create the row
              for (colNum = (short) 0; colNum < columnCount; colNum++) {
                  // cell = row.createCell(colNum);  // create the cell
                  int colType = rsmd.getColumnType (colNum+1); // get type

                  switch(colType) {

                     case Types.INTEGER:
                     case Types.SMALLINT:
                     case Types.BIGINT:
                     case Types.DOUBLE:
                     case Types.NUMERIC:
                     case Types.DECIMAL:
                      double dValue = rs.getDouble (colNum+1);
                      if (rs.wasNull ()) {
                        cell = row.createCell(colNum);
                        //cell.setCellValue(nullValue);
                        cell.setCellValue(emptyValue);
                      } else {
                        if ((dValue != 0) | (WRTZEROC.equals("Y"))) {
                          cell = row.createCell(colNum);  // create the cell
                          cell.setCellValue(dValue);
                        if ((colType == Types.DECIMAL) & (rsmd.getScale(colNum+1) == 0))
                          cell.setCellStyle(style0d);
                        if ((colType == Types.DECIMAL) & (rsmd.getScale(colNum+1) == 1))
                          cell.setCellStyle(style1d);    
                        if ((colType == Types.DECIMAL) & (rsmd.getScale(colNum+1) == 2))
                            cell.setCellStyle(style2d);
                        if ((colType == Types.DECIMAL) & (rsmd.getScale(colNum+1) == 3))
                            cell.setCellStyle(style3d);
                        if ((colType == Types.DECIMAL) & (rsmd.getScale(colNum+1) == 4))
                            cell.setCellStyle(style4d);
                        }
                      }
                      break;

                   case Types.DATE:
                      Date dateValue = rs.getDate (colNum+1);
                      if (rs.wasNull ()) {
                        cell = row.createCell(colNum);  // create the cell
                        //cell.setCellValue(nullValue);
                        cell.setCellValue(emptyValue);
                      } else {
                        cell = row.createCell(colNum);  // create the cell
                        cell.setCellValue(dateValue);
                        cell.setCellStyle(styleDate);
                      }
                      break;
                   case Types.TIME:
                      String timeValue = rs.getString (colNum+1);
                      if (rs.wasNull ()) {
                        cell = row.createCell(colNum);  // create the cell
                        //cell.setCellValue(nullValue);
                        cell.setCellValue(emptyValue);
                      } else {
                        cell = row.createCell(colNum);  // create the cell
                        cell.setCellValue(timeValue);
                        cell.setCellStyle(styleTime);
                      }
                     break;

                   case Types.CHAR:
                   default:
                      String aValue = rs.getString (colNum+1);
                      if (rs.wasNull ()) aValue = "<null>";
                      else {
                        if (LTRIM.equals("Y"))
                          aValue = aValue.trim();
                        else
                          if (aValue.trim().length() > 0) aValue = trimRight(aValue);
                          else break;
                        if (aValue.length() > 0) {
                          cell = row.createCell(colNum);  // create the cell
                          cell.setCellValue(aValue);
                        }
                      }
                  }

              }
            }
            // Make sheet HIDDEN or PROTECTED
            for (int i=0; i<wb.getNumberOfSheets(); i++)
            { String sName = wb.getSheetName(i); 
              if (sName.endsWith("_PH") || sName.endsWith("_HP")){
                wb.setSheetHidden(i,true);
                wb.getSheetAt(i).protectSheet("Udo");
              }else if (sName.endsWith("_P")){
                wb.getSheetAt(i).protectSheet("Udo");
              }else if (sName.endsWith("_H")){
                wb.setSheetHidden(i,true);
              }
             
            }


            // set the Column Width
		 	      if ((FROMXLS.equals(" ")) || (Integer.valueOf(BLANKROWS).intValue() < 0)) {
              if (rowNum > 0) {
                for (colNum = (short) 0; colNum < columnCount; colNum++) {
                  double units = (double) 1 / 256;
                  double pixels = ((columnWitdh[colNum] *1.1 ) / units) + 300;
                  columnWitdh[colNum] = (int) pixels;
                  if (columnWitdh[colNum] > 30000) columnWitdh[colNum] = 30000;
                  sheet.setColumnWidth(colNum, (short) columnWitdh[colNum]);
                }
              }
            }

            // show elapsed
            if (DEBUG.equals("Y")) {
              Date dateEnded = new Date();
              long timeEnded = dateEnded.getTime();
              long elapsed = Math.abs(timeEnded - timeStarted) / 1000;
              calcHMS(elapsed);
            }

            // write the Workbook
            if (DEBUG.equals("Y")) {
              System.out.println("Writing " + TOXLS);
              System.out.println(rowNum + " Rows x " + colNum + " Columns");
            }
            FileOutputStream fileOut = new FileOutputStream(TOXLS);
            wb.write(fileOut);
            fileOut.close();

            // show elapsed
            if (DEBUG.equals("Y")) {
              Date dateEnded = new Date();
              long timeEnded = dateEnded.getTime();
              long elapsed = Math.abs(timeEnded - timeStarted) / 1000;
              calcHMS(elapsed);
            }

        }

        catch (Exception e) {
            System.out.println ();
            System.out.println ("ERROR: " + e.getMessage());
            System.exit(1);
        }

        finally {

            // Clean up.
            try {
                if (connection != null)
                    connection.close ();
            }
            catch (SQLException e) {
                // Ignore.
            }
        }

      }  catch (Exception e) {
         System.out.println("Impossible to read SQL parameter");
         System.out.println(e.getMessage());
         System.exit(1);
      }
    } else {
       System.out.println("");
       System.out.println("The parameters are not correct.");
       System.out.println("");
       System.exit(1);
    }

  System.exit(0);

  }

  // -------------------------------------------------------------------------------
  // set columh header
  // -------------------------------------------------------------------------------
  public static ColHeader setColHeader(
    final int colNum,
    final String aInput,
    final String aWrapMarker,
    final String aNewLine) {
      ColHeader hdr = new ColHeader();
      final StringBuffer result = new StringBuffer();
      int startIdx = 0;
      int idxOld = 0;
      int maxIdx = 0;
      boolean wrap = false;
      while ((idxOld = aInput.indexOf(aWrapMarker, startIdx)) >= 0) {

        if (startIdx == 0)
          maxIdx = idxOld;
        else {
          int wrkIdx = idxOld - startIdx;
          if (wrkIdx > maxIdx) maxIdx = wrkIdx;
        }

        result.append( aInput.substring(startIdx, idxOld) );
        result.append( aNewLine );
        startIdx = idxOld + aWrapMarker.length();
      }
      result.append(aInput.substring(startIdx));

      if (maxIdx > 0) {
        int wrkIdx = aInput.length() - startIdx;
        if (wrkIdx > maxIdx) maxIdx = wrkIdx;
        wrap = true;
      }
      else {
        maxIdx = aInput.length();
      }

      hdr.label = result.toString();
      hdr.width = maxIdx + 1;
      hdr.wrap = wrap;
      return hdr;
  }

  // -------------------------------------------------------------------------------
  // copy SQL statement
  // -------------------------------------------------------------------------------
  public static void copySQL(InputStream in, OutputStream out)
   throws IOException {

    synchronized (in) {
      synchronized (out) {

        byte[] buffer = new byte[256];
        while (true) {
          int bytesRead = in.read(buffer);
          if (bytesRead == -1) break;
          out.write(buffer, 0, bytesRead);
        }
      }
    }
  }

  // -------------------------------------------------------------------------------
  // trim string Right
  // -------------------------------------------------------------------------------
  public static String trimRight(String s) {
     String t = s;
     while(t.charAt(t.length()-1) == ' ') {
        t = t.substring(0, t.length()-1);
      }
     return t;
    }

  // -------------------------------------------------------------------------------
  // show elapsed
  // -------------------------------------------------------------------------------
  public static void calcHMS(long timeInSeconds) {
      long hours, minutes, seconds;
      hours = timeInSeconds / 3600;
      timeInSeconds = timeInSeconds - (hours * 3600);
      minutes = timeInSeconds / 60;
      timeInSeconds = timeInSeconds - (minutes * 60);
      seconds = timeInSeconds;
      System.out.println(hours + " hour(s) " +
      minutes + " minute(s) " +
      seconds + " second(s)");
  }
}