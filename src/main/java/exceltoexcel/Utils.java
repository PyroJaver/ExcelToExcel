package exceltoexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class Utils {
    public HashMap<String, String> prepareProperties(String pathToProperties) throws IOException {

        //получаем доступ к файлу пропертис с именами анализов
        String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
        String appConfigPath = rootPath + pathToProperties;
        Properties readingProperties = new Properties();
        readingProperties.load(new FileInputStream(appConfigPath));
        HashMap<String, String> readyProperties = new HashMap<>();
        //считываем все записи из пропертис и заполняем хэшмапу с типами испытаний
        for (Map.Entry<Object, Object> propertiesEntrySet : readingProperties.entrySet()) {
            readyProperties.put((String) propertiesEntrySet.getKey(), (String) propertiesEntrySet.getValue());
        }
        return readyProperties;
    }
}
