package fileDate.file;

import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.csvreader.CsvReader;

import fileDate.model.Infos;

public class ReadFile {

	private LinkedHashMap<String, Map<String, Infos>> fileData = new LinkedHashMap<String, Map<String, Infos>>();
	private LinkedHashMap<String, List<Infos>> fileData2 = new LinkedHashMap<String, List<Infos>>();

	public LinkedHashMap<String, List<Infos>> readCSV() {
		try {

			List<Infos> infosList = new ArrayList<>();
			LinkedHashMap<String, Infos> fileData1 = new LinkedHashMap<String, Infos>();

			// 定义一个CSV路径
			String csvFilePath = "D://tt.csv";
			// 创建CSV读对象 例如:CsvReader(文件路径，分隔符，编码格式);
			CsvReader reader = new CsvReader(csvFilePath, ',', Charset.forName("UTF-8"));
			// 跳过表头 如果需要表头的话，这句可以忽略
			reader.readHeaders();
			// 逐行读入除表头的数据
			while (reader.readRecord()) {
				String[] aa = reader.getRawRecord().split(",");
				Infos infos = new Infos();
				infos.setmNum(aa[0]);
				infos.settNum(aa[1]);
				infos.setpName(aa[2]);
				infos.setpMoney(Double.parseDouble(aa[3]));
				infos.setPayType(aa[4]);
				infos.setDate(aa[5].split(" ")[0].replace('"', ' ').trim());
				infosList.add(infos);
			}
			reader.close();

			for (Infos infos : infosList) {
				fileData1.put(infos.gettNum(), infos);
				fileData.put(infos.getmNum(), fileData1);
			}

			Iterator<Entry<String, Map<String, Infos>>> it = fileData.entrySet().iterator();
			while (it.hasNext()) {
				Entry<String, Map<String, Infos>> entry = it.next();
				Iterator<Entry<String, Infos>> it2 = entry.getValue().entrySet().iterator();

				List<Infos> infosList2 = new ArrayList<>();
				while (it2.hasNext()) {
					Entry<String, Infos> entry2 = it2.next();
					if (entry2.getValue().getmNum().equals(entry.getKey())) {

						infosList2.add(entry2.getValue());
					}
				}
				fileData2.put(entry.getKey(), infosList2);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		return fileData2;
	}

}
