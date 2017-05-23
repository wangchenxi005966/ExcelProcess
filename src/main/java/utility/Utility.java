package utility;

import gui.MainFrame;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;

public class Utility {
    private static final int ARRAY_SIZE = 256;
    private Workbook climateWorkbook;
    private String climateFilePath;
    private CellStyle cellStyle;
    private int currentIndex = 0; //标识当前插入的行号
    private static final Logger logger = Logger.getLogger(Utility.class);
    private String[] headers = {"时间", "太阳辐射值", "No1", "No2", "No3", "风向", "气压", "相对湿度", "温度", "风速", "雨量", "小时累计雨量"};
    private int[] dataIndex = new int[ARRAY_SIZE];
    private String[] time = new String[ARRAY_SIZE];
    private double[] NO1 = new double[ARRAY_SIZE];
    private double[] NO2 = new double[ARRAY_SIZE];
    private double[] NO3 = new double[ARRAY_SIZE];
    private double[] direction = new double[ARRAY_SIZE];
    private double[] pressure = new double[ARRAY_SIZE];
    private double[] humity = new double[ARRAY_SIZE];
    private double[] temperature = new double[ARRAY_SIZE];
    private double[] speed = new double[ARRAY_SIZE];
    private double[] rainfall = new double[ARRAY_SIZE];
    private int[] sheetNumbers = new int[ARRAY_SIZE];
    private int count = 0;
    private int startSheetNumber = 0;
    private int startRowNumber = 1;

    public Utility(Workbook climate, String climateFilePath) {
        this.climateWorkbook = climate;
        this.climateFilePath = climateFilePath;
        this.cellStyle = this.climateWorkbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(HSSFColor.RED.index);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        this.climateWorkbook.createSheet();
    }

    public void query(double radiation, Date start, Date end, int index, MainFrame mainFrame, String traceFileName) {
        int[] result = this.getIndexInfo(start, end);//获取开始时间和结束时间的标号信息
        int startIndex = result[0];
        int endIndex = result[1];
        int currentStartSheetIndex = result[2];
        int currentEndSheetIndex = result[3];
        logger.info("开始编号:" + (startIndex + 1) + "; 表名:" + this.climateWorkbook.getSheetAt(currentStartSheetIndex).getSheetName()
                + "\n" + "结束编号:" + (endIndex + 1) + "; 表名:" + this.climateWorkbook.getSheetAt(currentEndSheetIndex).getSheetName());
        if (startIndex != 0 && endIndex != 0) {
            int a = 1;
            if (currentStartSheetIndex == currentEndSheetIndex) {//开始编号和结束编号位于同一个sheet
                Sheet currentSheet = this.climateWorkbook.getSheetAt(currentStartSheetIndex);
                for (int i = startIndex; i <= endIndex; i++) {
                    Row row = currentSheet.getRow(i);
                    double value = row.getCell(1).getNumericCellValue();
                    if (value == radiation) {
                        dataIndex[count] = i;
                        sheetNumbers[count] = currentStartSheetIndex;
                        this.setValue(row, index, a);
                        a++;
                    }
                }
            } else {
                Sheet startSheet = this.climateWorkbook.getSheetAt(currentStartSheetIndex);
                for (int i = startIndex; i <= startSheet.getLastRowNum(); i++) {
                    Row row = startSheet.getRow(i);
                    double value = row.getCell(1).getNumericCellValue();
                    if (value == radiation) {
                        dataIndex[count] = i;
                        sheetNumbers[count] = currentStartSheetIndex;
                        this.setValue(row, index, a);
                        a++;
                    }
                }
                Sheet endSheet = this.climateWorkbook.getSheetAt(currentEndSheetIndex);
                for (int i = 1; i <= endIndex; i++) {
                    Row row = endSheet.getRow(i);
                    double value = row.getCell(1).getNumericCellValue();
                    if (value == radiation) {
                        dataIndex[count] = i;
                        sheetNumbers[count] = currentEndSheetIndex;
                        this.setValue(row, index, a);
                        a++;
                    }
                }
            }
            if (count == 0) {
                this.appendErrorMessage(mainFrame, traceFileName, "未找到相应的跟踪数据值出现行", index, radiation, start, end);
            } else {
                this.writeData(radiation);
            }
        } else {
            this.appendErrorMessage(mainFrame, traceFileName, "未找到开始时间和结束时间对应的行编号", index, radiation, start, end);
        }
    }

    private void setValue(Row row, int index, int a) {
        time[count] = row.getCell(0).getStringCellValue();
        NO1[count] = row.getCell(2).getNumericCellValue();
        NO2[count] = row.getCell(3).getNumericCellValue();
        NO3[count] = row.getCell(4).getNumericCellValue();
        direction[count] = row.getCell(5).getNumericCellValue();
        pressure[count] = row.getCell(6).getNumericCellValue();
        humity[count] = row.getCell(7).getNumericCellValue();
        temperature[count] = row.getCell(8).getNumericCellValue();
        speed[count] = row.getCell(9).getNumericCellValue();
        rainfall[count] = row.getCell(10).getNumericCellValue();
        Cell cell = row.createCell(12);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(index + "-" + a);
        count++;
    }

    private void writeData(double radiation) {
        int sheetNumber = this.climateWorkbook.getNumberOfSheets();
        Sheet sheet = this.climateWorkbook.getSheetAt(sheetNumber - 1);
        if (sheet.getPhysicalNumberOfRows() == 0 && sheet.getLastRowNum() == 0) {//创建标题栏
            Row row = sheet.createRow(0);
            for (int i = 0; i < this.headers.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(headers[i]);
            }
        }
        currentIndex++;
        Row current = sheet.createRow(currentIndex);
        current.createCell(0);
        current.createCell(1).setCellValue(radiation);
        current.createCell(2).setCellValue(average(Arrays.copyOf(NO1, count)));
        current.createCell(3).setCellValue(average(Arrays.copyOf(NO2, count)));
        current.createCell(4).setCellValue(average(Arrays.copyOf(NO3, count)));
        current.createCell(5).setCellValue(average(Arrays.copyOf(direction, count)));
        current.createCell(6).setCellValue(average(Arrays.copyOf(pressure, count)));
        current.createCell(7).setCellValue(average(Arrays.copyOf(humity, count)));
        current.createCell(8).setCellValue(average(Arrays.copyOf(temperature, count)));
        current.createCell(9).setCellValue(average(Arrays.copyOf(speed, count)));
        current.createCell(10).setCellValue(average(Arrays.copyOf(rainfall, count)));
        current.createCell(11);

        double median = Math.ceil((count - 1) * 0.5);
        //set time
        String finalTime = time[(int) median];
        int index = dataIndex[(int) median];
        int currentSheetNumber = sheetNumbers[(int) median];
        int counter = 0;
        //calculate total rainfall
        double sumRainfall = 0;
        Sheet currentSheet = this.climateWorkbook.getSheetAt(currentSheetNumber);
        for (int j = index; j > 0; j--) {
            double value = currentSheet.getRow(j).getCell(10).getNumericCellValue();
            sumRainfall += value;
            counter++;
            if (counter == 1800) {
                break;
            }
        }
        current.getCell(0).setCellValue(finalTime);
        current.getCell(11).setCellValue(sumRainfall);
        count = 0;
    }

    private double average(double[] data) {
        double sum = 0;
        int length = data.length;
        for (double i : data) {
            sum += i;
        }
        return this.round(sum / length);
    }

    public Date parse(String dateString) {
        try {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            return sdf.parse(dateString);
        } catch (ParseException e) {
            logger.error("日期转化失败");
            e.printStackTrace();
            return null;
        }
    }

    public Date subtractOneCycle(Date date) {
        if (null == date) {
            return null;
        }
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.SECOND, -73);
        return calendar.getTime();
    }

    public void close() {
        try {
            Sheet sheet = this.climateWorkbook.getSheetAt(this.climateWorkbook.getNumberOfSheets() - 1);
            sheet.autoSizeColumn(0);
            for (int i = 1; i < headers.length; i++) {
                sheet.setColumnWidth(i, 12 * 256);
            }
            OutputStream outputStream = new FileOutputStream(this.climateFilePath);
            this.climateWorkbook.write(outputStream);
            outputStream.close();
            this.climateWorkbook.close();
        } catch (IOException e) {
            logger.error("气象文件写入失败");
            e.printStackTrace();
        }
    }

    private Date addOneSecond(Date date) {
        if (null == date) {
            return null;
        }
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.SECOND, 1);
        return calendar.getTime();
    }

    private boolean sameDate(Date d1, Date d2) {
        if (null == d1 || null == d2) {
            return false;
        }
        Calendar cal1 = Calendar.getInstance();
        cal1.setTime(d1);
        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(d2);
        int result = cal1.compareTo(cal2);
        return result == 0;
    }

    private String format(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return sdf.format(date);
    }

    private double round(double value) {
        BigDecimal bigDecimal = new BigDecimal(value);
        return bigDecimal.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
    }

    private void appendErrorMessage(MainFrame frame, String trace, String reason, int index, double radiation, Date start, Date end) {
        frame.appendText("******************************************************************************************");
        frame.appendText("跟踪数据失败:" + trace);
        frame.appendText("失败原因: " + reason);
        frame.appendText("数据编号:" + (index + 1));
        frame.appendText("辐射值:" + radiation);
        frame.appendText("开始时间:" + this.format(start));
        frame.appendText("结束时间:" + this.format(end));
        frame.appendText("******************************************************************************************");
        frame.write("******************************************************************************************");
        frame.write("跟踪数据失败:" + trace);
        frame.write("失败原因: " + reason);
        frame.write("数据编号:" + (index + 1));
        frame.write("辐射值:" + radiation);
        frame.write("开始时间:" + this.format(start));
        frame.write("结束时间:" + this.format(end));
        frame.write("******************************************************************************************");
    }

    private int[] getIndexInfo(Date start, Date end) {
        int[] result = new int[4];
        int sheetNumber = this.climateWorkbook.getNumberOfSheets();
        outer:
        for (int j = startSheetNumber; j < sheetNumber - 1; j++) {
            Sheet sheet = this.climateWorkbook.getSheetAt(j);
            if (!sheet.getSheetName().contains("data")) {
                logger.error("遍历到非数据页");
                break;
            }
            int rowNumber = sheet.getLastRowNum();
            for (int i = startRowNumber; i < rowNumber; i++) {
                Row row = sheet.getRow(i);
                String dateString = row.getCell(0).getStringCellValue();
                Date date = this.parse(dateString);
                if (this.sameDate(start, date) || this.sameDate(start, this.addOneSecond(date)) ||
                        this.sameDate(this.addOneSecond(start), date)) {
                    result[0] = (i - 1 > 0) ? (i - 1) : i;
                    result[2] = j;
                    continue;
                }
                if (this.sameDate(end, date) || this.sameDate(this.addOneSecond(end), date) ||
                        this.sameDate(end, this.addOneSecond(date))) {
                    result[1] = i + 1;
                    result[3] = j;
                    this.startSheetNumber = j;
                    this.startRowNumber = i;
                    break outer;
                }
            }
            //执行跳页操作
            startRowNumber = 1;
        }
        return result;
    }
}
