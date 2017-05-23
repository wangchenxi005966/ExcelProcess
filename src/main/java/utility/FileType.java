package utility;

import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.HttpHeaders;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.metadata.TikaMetadataKeys;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class FileType {
    private final static HashMap<String, String> FILE_TYPE_MAP = new HashMap<>();

    static {
        getAllFileType();  //初始化文件类型信息
    }

    private static void getAllFileType() {
        FILE_TYPE_MAP.put("doc", "application/msword");
        FILE_TYPE_MAP.put("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        FILE_TYPE_MAP.put("xls", "application/vnd.ms-excel");
        FILE_TYPE_MAP.put("xlsx", "vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        FILE_TYPE_MAP.put("ppt", "application/vnd.ms-powerpoint");
        FILE_TYPE_MAP.put("pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        FILE_TYPE_MAP.put("rar", "application/x-rar-compressed");
        FILE_TYPE_MAP.put("zip", "application/zip");
        FILE_TYPE_MAP.put("pdf", "application/pdf");
        FILE_TYPE_MAP.put("exe", "application/octet-stream");
        FILE_TYPE_MAP.put("avi", "video/x-msvideo");
        FILE_TYPE_MAP.put("bmp", "image/bmp");
        FILE_TYPE_MAP.put("jpg", "image/jpeg");
        FILE_TYPE_MAP.put("gif", "image/gif");
        FILE_TYPE_MAP.put("txt", "text/plain");
        FILE_TYPE_MAP.put("css", "text/css");
        FILE_TYPE_MAP.put("html", "text/html");
        FILE_TYPE_MAP.put("java", "text/x-java-source");
        FILE_TYPE_MAP.put("c", "text/x-csrc");
        FILE_TYPE_MAP.put("c++", "text/x-c++src");
    }

    public String getMimeType(File file) {
        if (file.isDirectory()) {
            return "the target is a directory";
        }
        AutoDetectParser parser = new AutoDetectParser();
        parser.setParsers(new HashMap<>());
        Metadata metadata = new Metadata();
        metadata.add(TikaMetadataKeys.RESOURCE_NAME_KEY, file.getName());
        InputStream stream;
        try {
            stream = new FileInputStream(file);
            parser.parse(stream, new DefaultHandler(), metadata, new ParseContext());
            stream.close();
        } catch (TikaException | SAXException | IOException e) {
            e.printStackTrace();
        }
        String fileType = metadata.get(HttpHeaders.CONTENT_TYPE);
        for (Map.Entry<String, String> entry : FILE_TYPE_MAP.entrySet()) {
            String fileTypeValue = entry.getValue();
            if (fileType.equals(fileTypeValue)) {
                return entry.getKey();
            }
        }
        return null;
    }
}
