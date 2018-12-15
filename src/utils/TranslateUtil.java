package utils;

import org.dom4j.DocumentException;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * author: jixiaoyong
 * email: jixiaoyong1995@gmail.com
 * website: www.jixiaoyong.github.io
 * date: 2018/12/15
 * description: todo
 */
public class TranslateUtil {

    public static Map<String, Map<String, Object>> readXmls(String path) {
        Map<String, Map<String, Object>> maps = new HashMap<>();
        File file = new File(path);
        File[] xmlFiles = file.listFiles();
        for (int i = 0; i < xmlFiles.length; i++) {
            if (!xmlFiles[i].getName().startsWith("values")) {
                continue;
            }
            try {
                String language = xmlFiles[i].getAbsolutePath().substring(xmlFiles[i].getAbsolutePath().lastIndexOf(File.separator) + 1)
                        .replace("values", "");
                if ("".equals(language)) {
                    language = "default";
                }
                Map<String, Object> map = null;
                try {
                    map = XMLUtil.readFormatXML(new File(xmlFiles[i].getAbsolutePath() + File.separator + "strings.xml"));
                } catch (DocumentException e) {
                    map = XMLUtil.readFormatXML(new File(xmlFiles[i].getAbsolutePath() + File.separator + "string.xml"));
                    e.printStackTrace();
                }
                maps.put(language, map);
            } catch (DocumentException e) {
                e.printStackTrace();
            }
        }
        return maps;
    }

    public static Map<String, Map<String, Object>> readExcel(String excelPath, String outputDirPath) {
        return ExcelUtil.readExcel(excelPath, outputDirPath);
    }


    public static void writeToExcel(String outputFilePath, Map<String, Map<String, Object>> maps) {
        Set<String> mapKetSet = maps.keySet();
        ExcelUtil.writXLSXExcel("default", 1, maps.get("default"), outputFilePath);
        maps.remove("default");
        for (int i = 0; i < mapKetSet.size(); i++) {
            ExcelUtil.writXLSXExcel((String) mapKetSet.toArray()[i], i + 2, maps.get(mapKetSet.toArray()[i]), outputFilePath);
        }
    }

    public static void writeToXML(String path, Map<String, Map<String, Object>> maps) {
        Set<String> mapKetSet = maps.keySet();

        ExcelUtil.writXLSXExcel("default", 1, maps.get("default"), path);
        maps.remove("default");
        for (int i = 0; i < mapKetSet.size(); i++) {
            ExcelUtil.writXLSXExcel((String) mapKetSet.toArray()[i], i + 2, maps.get(mapKetSet.toArray()[i]), path);
        }
    }
}
