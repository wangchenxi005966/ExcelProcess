package gui;

import entity.ExcelWorkBook;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import utility.FileType;
import utility.Utility;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.Properties;

public class MainFrame extends JFrame {
    private JTextArea textArea = null;
    private static final Logger logger = Logger.getLogger(MainFrame.class);
    private ArrayList<File> totalFiles = new ArrayList<>();
    private ArrayList<File> traceFiles = new ArrayList<>();
    private ArrayList<File> climateFiles = new ArrayList<>();
    private PrintWriter printWriter;
    private String directoryName = "Result";

    public MainFrame() {
        Container contentPane = this.getContentPane();
        contentPane.setLayout(new BorderLayout());
        contentPane.add(getScrollPanel(), BorderLayout.CENTER);
        contentPane.add(getPanel(), BorderLayout.SOUTH);
        this.setTitle("文件窗口");
        this.setSize(1200, 600);
        Toolkit toolkit = Toolkit.getDefaultToolkit();
        int x = (int) (toolkit.getScreenSize().getWidth() - this.getWidth()) / 2;
        int y = (int) (toolkit.getScreenSize().getHeight() - this.getHeight()) / 2;
        this.setLocation(x, y);
        this.setResizable(false);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setVisible(true);
        this.initLog4j();
    }

    private Container getScrollPanel() {
        textArea = new JTextArea(10, 15);
        textArea.setTabSize(4);
        textArea.setFont(new Font("微软雅黑", Font.PLAIN, 20));
        textArea.setLineWrap(true);// 激活自动换行功能
        textArea.setWrapStyleWord(true);// 激活断行不断字功能
        textArea.setEditable(false);
        return new JScrollPane(textArea);
    }

    private Container getPanel() {
        JPanel panel = new JPanel();

        JButton executeButton = new JButton("开始执行");
        executeButton.setBorderPainted(false);
        executeButton.setFont(new Font("微软雅黑", Font.PLAIN, 24));
        executeButton.addActionListener(e -> {
            Thread thread = new Thread(() -> {
                JFileChooser fileChooser = new JFileChooser("E:\\");
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int option = fileChooser.showDialog(new JLabel(), "请选择数据目录");
                if (option == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    appendText("已选择目录: " + file.getAbsolutePath());
                    if (this.generateResultDirectory()) {
                        this.traverseFolder(file);
                    } else {
                        this.appendText("创建日志文件目录失败");
                    }
                }
            });
            thread.start();
        });
        panel.add(executeButton);
        return panel;
    }

    public void appendText(String text) {
        textArea.append(text);
        textArea.append("\n");
        textArea.setCaretPosition(textArea.getText().length());
    }

    private void traverseFolder(File directory) {
        FileType fileType = new FileType();
        this.totalFiles.clear();
        if (directory.exists()) {
            File[] files = directory.listFiles();
            if ((files != null ? files.length : 0) == 0) {
                this.appendText("文件夹是空的!");
            } else {
                for (File file : files) {
                    if (file.isFile()) {
                        this.appendText("文件:" + file.getAbsolutePath());
                        String type = fileType.getMimeType(file);
                        if ("xls".equals(type) || "xlsx".equals(type)) {
                            this.totalFiles.add(file);
                        }
                    }
                }
            }
        } else {
            this.appendText("文件不存在!");
        }
        this.distributeFiles();
    }

    private void distributeFiles() {
        for (File file : this.totalFiles) {
            if (file.getName().contains("跟踪数据")) {
                this.traceFiles.add(file);
            } else if (file.getName().contains("气象数据")) {
                this.climateFiles.add(file);
            }
        }
        this.appendText("数据写入中……");
        this.getFileName();
    }

    private void getFileName() {
        long totalTime = 0;
        int totalFileNumber = this.traceFiles.size();
        int current = 1;
        try {
            for (File traceFile : this.traceFiles) {
                this.appendText("正在处理第" + current + "/" + totalFileNumber + "个文件……");
                String traceFileName = traceFile.getName();
                String prefix = traceFileName.substring(0, traceFileName.indexOf("跟踪数据"));//获取文件名前缀
                for (File climateFile : this.climateFiles) {
                    String climateFileName = climateFile.getName();
                    if (climateFileName.contains(prefix)) {
                        this.generateResultFile(prefix);
                        long cost = this.execute(traceFile.getAbsolutePath(),
                                climateFile.getAbsolutePath(), climateFileName, traceFileName);
                        totalTime += cost;
                        this.close();
                    }
                }
                current++;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        this.appendText("总时间:" + totalTime + "s");
    }

    private long execute(String tracePath, String climatePath, String climateFileName, String traceFileName) throws IOException {
        Long startTime = System.currentTimeMillis();
        ExcelWorkBook excelWorkBook = new ExcelWorkBook();
        Workbook trace = excelWorkBook.getExcelWorkBook(tracePath);
        Workbook climate = excelWorkBook.getExcelWorkBook(climatePath);
        Utility utility = new Utility(climate, climatePath);
        Sheet sheet = trace.getSheetAt(0);
        int rowNumber = sheet.getLastRowNum();
        int index = 2;

        for (int i = 2; i <= rowNumber; i++) {
            Row rowCurrent = sheet.getRow(i);
            Row rowBefore = sheet.getRow(i - 1);
            //获取辐射值
            double radiation = rowCurrent.getCell(2).getNumericCellValue();
            //获取起止时间
            String start = rowBefore.getCell(5).getStringCellValue();
            String end = rowCurrent.getCell(5).getStringCellValue();

            Date startDate = utility.parse(start);
            Date endDate = utility.parse(end);

            if (startDate != null && endDate != null) {
                logger.info("radiation:" + radiation + "; start:" + start + "; end:" + end);
                if ((endDate.getTime() - startDate.getTime()) / 1000 > 10000) {//判断时期时间是否跨天
                    utility.query(radiation, utility.subtractOneCycle(endDate), endDate, index, this, traceFileName);
                } else {
                    utility.query(radiation, startDate, endDate, index, this, traceFileName);
                }
                index++;
            }
        }
        utility.close();
        trace.close();
        Long endTime = System.currentTimeMillis();
        this.appendText(climateFileName + "文件数据写入完毕");
        long cost = (endTime - startTime) / 1000;
        this.appendText("花费时间：" + cost + "s");
        return cost;
    }

    private void initLog4j() {
        Properties properties = new Properties();
        InputStream inputStream = this.getClass().getResourceAsStream("/log4j.properties");
        try {
            properties.load(inputStream);
            PropertyConfigurator.configure(properties);
            logger.info("log4j自定义配置文件初始完毕");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private boolean generateResultDirectory() {
        File directory = new File(this.directoryName);
        if (!directory.exists()) {
            if (!directory.mkdir()) {
                logger.error("创建目录失败");
                return false;
            }
        }
        return true;
    }

    private void generateResultFile(String name) {
        try {
            this.printWriter = new PrintWriter(this.directoryName + "/" + name + "日志记录.txt", "UTF-8");
        } catch (FileNotFoundException | UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }

    private void close() {
        this.printWriter.close();
    }

    public void write(String message) {
        this.printWriter.println(message);
    }
}
