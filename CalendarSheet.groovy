@Grab('org.apache.logging.log4j:log4j-core:2.19.0')
@Grab('org.apache.poi:poi-ooxml:4.1.2')

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.time.DayOfWeek
import java.time.LocalDate
import java.time.Month
import java.time.YearMonth

enum CellRole {
    TITLE,
    MONTH_LABEL,
    DAY_LABEL,
    WEEKDAY,
    WEEKEND,
    INDEX,
    PADDING,
}

@groovy.transform.TupleConstructor
@groovy.transform.ToString
class CellInfo {
    Integer column
    Integer row
    String text
    CellRole role
}

def static main(String[] args) {
    if (args.length != 2 || !args[0].matches("\\d+") || !args[1].matches(".*\\.xlsx") ) {
        println "Syntax: groovy CalendarSheet.groovy <4-digit-year> <filename.xlsx>"
        return
    }
    def year = Integer.parseInt(args[0])
    def outputFile = new File(args[1])
    if (outputFile.exists()) {
        println "File '${outputFile.name}' already exists."
        return
    }
    run(year, outputFile)
}

static void run(int year, File outputFile) {
    println("Calculating calendar for ${year} ...")
    CellInfo[] cellInfos = calcCellInfos(year)
    CellInfo[][] rowsOfCells = asRowsOfCells(cellInfos)
    println("Writing calendar to ${outputFile.name} ...")
    generateSheet(rowsOfCells, outputFile)
    println("Done.")
}

static CellInfo[] calcCellInfos(int year) {
    def cellInfos = []
    (1..24).each { column ->
        def text = column == 2 ? "${year} Calendar" : ""
        cellInfos.add(new CellInfo(column, 0, text, CellRole.TITLE))
    }
    Month.iterator().eachWithIndex { def month, int column ->
        cellInfos.add(new CellInfo(column * 2 + 1, 1, "", CellRole.INDEX))
        cellInfos.add(new CellInfo(column * 2 + 2, 1, month.toString().substring(0, 3), CellRole.MONTH_LABEL))
    }
    Month.values().eachWithIndex { Month month, int column ->
        def dayCount = YearMonth.of(year, month).lengthOfMonth()
        def offset = LocalDate.of(year, month, 1).dayOfWeek.ordinal()
        (1..dayCount).forEach { dayNumber ->
            def weekend = LocalDate.of(year, column + 1, dayNumber).dayOfWeek in [DayOfWeek.SATURDAY, DayOfWeek.SUNDAY]
            def role = weekend ? CellRole.WEEKEND : CellRole.WEEKDAY
            cellInfos.add(new CellInfo(column * 2 + 1, 1 + offset + dayNumber, "${dayNumber}", CellRole.INDEX))
            cellInfos.add(new CellInfo(column * 2 + 2, 1 + offset + dayNumber, "", role))
        }
    }
    def maxRow = cellInfos*.row.max()
    (2..maxRow).iterator().each { row ->
        def day = DayOfWeek.values()[(row - 2) % DayOfWeek.values().size()]
                .toString().substring(0, 3).toLowerCase().capitalize()
        cellInfos.add(new CellInfo(0, row, day, CellRole.DAY_LABEL))
        cellInfos.add(new CellInfo(25, row, day, CellRole.DAY_LABEL))
    }
    return cellInfos
}

static CellInfo[][] asRowsOfCells(CellInfo[] cellInfos) {
    def maxColumn = cellInfos*.column.max()
    def maxRow = cellInfos*.row.max()
    def result = [null] * maxRow
    (0..maxRow).iterator().each { row ->
        result[row] = [null] * maxColumn
    }
    cellInfos.iterator().each {
        result[it.row][it.column] = it
    }
    (0..maxRow).iterator().each { row ->
        (0..maxColumn).iterator().each { column ->
            if (result[row][column] == null) {
                result[row][column] = new CellInfo(column, row, "", CellRole.PADDING)
            }
        }
    }
    return result
}

static <T> T[][] pivot(T[][] data) {
    def oldPrimary = data.length
    def oldSecondary = data*.length.max()
    def result = [null] * oldSecondary
    (0..oldSecondary - 1).iterator().each { primary ->
        result[primary] = [null] * oldPrimary
        (0..oldPrimary - 1).iterator().each { secondary ->
            result[primary][secondary] = data[secondary][primary]
        }
    }
    return result
}

static void generateSheet(CellInfo[][] rowsOfCells, File outputFile) {
    def final CHAR_POI_WIDTH_FACTOR = 256
    def workbook = new XSSFWorkbook()
    def sheet = workbook.createSheet("Calendar")
    rowsOfCells.eachWithIndex { CellInfo[] cellInfoRow, int row ->
        def sheetRow = sheet.createRow(row)
        cellInfoRow.eachWithIndex { CellInfo cellInfo, int column ->
            def cell = sheetRow.createCell(column)
            if (cellInfo.text.matches("\\d+")) {
                cell.setCellType(CellType.NUMERIC)
                cell.setCellValue(Integer.parseInt(cellInfo.text))
            } else {
                cell.setCellType(CellType.STRING)
                cell.setCellValue(cellInfo.text)
            }
            cell.setCellStyle(cellStyleForRole(cellInfo.role, cell))
        }
    }
    CellInfo[][] columnsOfCells = pivot(rowsOfCells)
    columnsOfCells.eachWithIndex { CellInfo[] columnCellInfos, int column ->
        def columnRole = columnCellInfos*.role.find {
            it in [CellRole.MONTH_LABEL, CellRole.DAY_LABEL, CellRole.INDEX]
        }
        sheet.setColumnWidth(column, columnWidthInCharsForRole(columnRole) * CHAR_POI_WIDTH_FACTOR)
    }
    new FileOutputStream(outputFile).withCloseable {
        workbook.write(it)
    }
}

static CellStyle cellStyleForRole(CellRole cellRole, XSSFCell cell) {
    CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
    switch (cellRole) {
        case CellRole.TITLE:
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.MONTH_LABEL:
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.DAY_LABEL:
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.WEEKDAY:
            cellStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.WEEKEND:
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.INDEX:
            cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            break
        case CellRole.PADDING:
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex())
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    }
    return cellStyle
}

static int columnWidthInCharsForRole(CellRole columnRole) {
    switch (columnRole) {
        case CellRole.DAY_LABEL:
            return 6
        case CellRole.INDEX:
            return 4
        case CellRole.MONTH_LABEL:
        default:
            return 14
    }
}
