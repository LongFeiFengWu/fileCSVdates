package fileDate.file;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import fileDate.model.Infos;

public class OutputFile {

	public static void main(String[] args) throws IOException {

		System.out.println("---请先保证D盘根目录下存在已导出的原始数据，并重命名为tt.csv---(Y/N)");

		int forSure = System.in.read();

		if (forSure == 89 || forSure == 121) {
			System.out.println("---开始整理数据---");
			long startTime = System.currentTimeMillis();// 记录开始时间

			OutputFile outputFile = new OutputFile();
			outputFile.getResult();
			long endTime = System.currentTimeMillis();// 记录结束时间
			float excTime = (float) (endTime - startTime) / 1000;

			System.out.println("---整理数据完成，结果保存在D盘根目录下 统计结果.xls，花费时间---：" + excTime + "s");

		} else {
			return;
		}

	}

	private void getResult() {
		ReadFile file = new ReadFile();
		LinkedHashMap<String, List<Infos>> fileData = file.readCSV();

		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();

		HSSFSheet sheet18640 = wb.createSheet("18640光耀东方总部");
		HSSFSheet sheet19452 = wb.createSheet("19452光耀东方广场");
		HSSFSheet sheet19453 = wb.createSheet("19453联动优势1");
		HSSFSheet sheet19454 = wb.createSheet("19454联动优势9");
		HSSFSheet sheet19455 = wb.createSheet("19455烽火科技");
		HSSFSheet sheet19456 = wb.createSheet("19456房天下3");
		HSSFSheet sheet19457 = wb.createSheet("19457房天下1");
		HSSFSheet sheet19458 = wb.createSheet("19458H3C");
		HSSFSheet sheet19459 = wb.createSheet("19459爱农驿站22");
		HSSFSheet sheet19472 = wb.createSheet("19472百融金服北楼20");
		HSSFSheet sheet19473 = wb.createSheet("19473百融金服南楼2");
		HSSFSheet sheet19475 = wb.createSheet("19475海象金服2");
		HSSFSheet sheet19481 = wb.createSheet("19481云海肴");
		HSSFSheet sheet19482 = wb.createSheet("19482意锐新创D5");
		HSSFSheet sheet19483 = wb.createSheet("19483中科金财");
		HSSFSheet sheet19484 = wb.createSheet("19484丰瑞祥");

		HSSFSheet sheetSum = wb.createSheet("汇总");

		HSSFRow row18640 = sheet18640.createRow((int) 0);
		HSSFCell cell18640 = row18640.createCell((short) 0);
		cell18640.setCellValue("日期");
		cell18640 = row18640.createCell((short) 1);
		cell18640.setCellValue("付费杯数");
		cell18640 = row18640.createCell((short) 2);
		cell18640.setCellValue("收费金额");
		cell18640 = row18640.createCell((short) 3);
		cell18640.setCellValue("免费杯数");
		cell18640 = row18640.createCell((short) 4);
		cell18640.setCellValue("免单额");
		cell18640 = row18640.createCell((short) 5);
		cell18640.setCellValue("美式");
		cell18640 = row18640.createCell((short) 6);
		cell18640.setCellValue("拿铁");
		cell18640 = row18640.createCell((short) 7);
		cell18640.setCellValue("摩卡");
		cell18640 = row18640.createCell((short) 8);
		cell18640.setCellValue("卡布奇诺");
		cell18640 = row18640.createCell((short) 9);
		cell18640.setCellValue("巧克力");
		cell18640 = row18640.createCell((short) 10);
		cell18640.setCellValue("玛琪雅朵");
		cell18640 = row18640.createCell((short) 11);
		cell18640.setCellValue("牛奶");
		cell18640 = row18640.createCell((short) 12);
		cell18640.setCellValue("巧克力牛奶");
		cell18640 = row18640.createCell((short) 13);
		cell18640.setCellValue("抹茶");
		cell18640 = row18640.createCell((short) 14);
		cell18640.setCellValue("抹茶咖啡");

		HSSFRow row19452 = sheet19452.createRow((int) 0);
		HSSFCell cell19452 = row19452.createCell((short) 0);
		cell19452.setCellValue("日期");
		cell19452 = row19452.createCell((short) 1);
		cell19452.setCellValue("付费杯数");
		cell19452 = row19452.createCell((short) 2);
		cell19452.setCellValue("收费金额");
		cell19452 = row19452.createCell((short) 3);
		cell19452.setCellValue("免费杯数");
		cell19452 = row19452.createCell((short) 4);
		cell19452.setCellValue("免单额");
		cell19452 = row19452.createCell((short) 5);
		cell19452.setCellValue("美式");
		cell19452 = row19452.createCell((short) 6);
		cell19452.setCellValue("拿铁");
		cell19452 = row19452.createCell((short) 7);
		cell19452.setCellValue("摩卡");
		cell19452 = row19452.createCell((short) 8);
		cell19452.setCellValue("卡布奇诺");
		cell19452 = row19452.createCell((short) 9);
		cell19452.setCellValue("巧克力");
		cell19452 = row19452.createCell((short) 10);
		cell19452.setCellValue("玛琪雅朵");
		cell19452 = row19452.createCell((short) 11);
		cell19452.setCellValue("牛奶");
		cell19452 = row19452.createCell((short) 12);
		cell19452.setCellValue("巧克力牛奶");
		cell19452 = row19452.createCell((short) 13);
		cell19452.setCellValue("抹茶");
		cell19452 = row19452.createCell((short) 14);
		cell19452.setCellValue("抹茶咖啡");

		HSSFRow row19453 = sheet19453.createRow((int) 0);
		HSSFCell cell19453 = row19453.createCell((short) 0);
		cell19453.setCellValue("日期");
		cell19453 = row19453.createCell((short) 1);
		cell19453.setCellValue("付费杯数");
		cell19453 = row19453.createCell((short) 2);
		cell19453.setCellValue("收费金额");
		cell19453 = row19453.createCell((short) 3);
		cell19453.setCellValue("免费杯数");
		cell19453 = row19453.createCell((short) 4);
		cell19453.setCellValue("免单额");
		cell19453 = row19453.createCell((short) 5);
		cell19453.setCellValue("美式");
		cell19453 = row19453.createCell((short) 6);
		cell19453.setCellValue("拿铁");
		cell19453 = row19453.createCell((short) 7);
		cell19453.setCellValue("摩卡");
		cell19453 = row19453.createCell((short) 8);
		cell19453.setCellValue("卡布奇诺");
		cell19453 = row19453.createCell((short) 9);
		cell19453.setCellValue("巧克力");
		cell19453 = row19453.createCell((short) 10);
		cell19453.setCellValue("玛琪雅朵");
		cell19453 = row19453.createCell((short) 11);
		cell19453.setCellValue("牛奶");
		cell19453 = row19453.createCell((short) 12);
		cell19453.setCellValue("巧克力牛奶");
		cell19453 = row19453.createCell((short) 13);
		cell19453.setCellValue("抹茶");
		cell19453 = row19453.createCell((short) 14);
		cell19453.setCellValue("抹茶咖啡");

		HSSFRow row19454 = sheet19454.createRow((int) 0);
		HSSFCell cell19454 = row19454.createCell((short) 0);
		cell19454.setCellValue("日期");
		cell19454 = row19454.createCell((short) 1);
		cell19454.setCellValue("付费杯数");
		cell19454 = row19454.createCell((short) 2);
		cell19454.setCellValue("收费金额");
		cell19454 = row19454.createCell((short) 3);
		cell19454.setCellValue("免费杯数");
		cell19454 = row19454.createCell((short) 4);
		cell19454.setCellValue("免单额");
		cell19454 = row19454.createCell((short) 5);
		cell19454.setCellValue("美式");
		cell19454 = row19454.createCell((short) 6);
		cell19454.setCellValue("拿铁");
		cell19454 = row19454.createCell((short) 7);
		cell19454.setCellValue("摩卡");
		cell19454 = row19454.createCell((short) 8);
		cell19454.setCellValue("卡布奇诺");
		cell19454 = row19454.createCell((short) 9);
		cell19454.setCellValue("巧克力");
		cell19454 = row19454.createCell((short) 10);
		cell19454.setCellValue("玛琪雅朵");
		cell19454 = row19454.createCell((short) 11);
		cell19454.setCellValue("牛奶");
		cell19454 = row19454.createCell((short) 12);
		cell19454.setCellValue("巧克力牛奶");
		cell19454 = row19454.createCell((short) 13);
		cell19454.setCellValue("抹茶");
		cell19454 = row19454.createCell((short) 14);
		cell19454.setCellValue("抹茶咖啡");

		HSSFRow row19455 = sheet19455.createRow((int) 0);
		HSSFCell cell19455 = row19455.createCell((short) 0);
		cell19455.setCellValue("日期");
		cell19455 = row19455.createCell((short) 1);
		cell19455.setCellValue("付费杯数");
		cell19455 = row19455.createCell((short) 2);
		cell19455.setCellValue("收费金额");
		cell19455 = row19455.createCell((short) 3);
		cell19455.setCellValue("免费杯数");
		cell19455 = row19455.createCell((short) 4);
		cell19455.setCellValue("免单额");
		cell19455 = row19455.createCell((short) 5);
		cell19455.setCellValue("美式");
		cell19455 = row19455.createCell((short) 6);
		cell19455.setCellValue("拿铁");
		cell19455 = row19455.createCell((short) 7);
		cell19455.setCellValue("摩卡");
		cell19455 = row19455.createCell((short) 8);
		cell19455.setCellValue("卡布奇诺");
		cell19455 = row19455.createCell((short) 9);
		cell19455.setCellValue("巧克力");
		cell19455 = row19455.createCell((short) 10);
		cell19455.setCellValue("玛琪雅朵");
		cell19455 = row19455.createCell((short) 11);
		cell19455.setCellValue("牛奶");
		cell19455 = row19455.createCell((short) 12);
		cell19455.setCellValue("巧克力牛奶");
		cell19455 = row19455.createCell((short) 13);
		cell19455.setCellValue("抹茶");
		cell19455 = row19455.createCell((short) 14);
		cell19455.setCellValue("抹茶咖啡");

		HSSFRow row19456 = sheet19456.createRow((int) 0);
		HSSFCell cell19456 = row19456.createCell((short) 0);
		cell19456.setCellValue("日期");
		cell19456 = row19456.createCell((short) 1);
		cell19456.setCellValue("付费杯数");
		cell19456 = row19456.createCell((short) 2);
		cell19456.setCellValue("收费金额");
		cell19456 = row19456.createCell((short) 3);
		cell19456.setCellValue("免费杯数");
		cell19456 = row19456.createCell((short) 4);
		cell19456.setCellValue("免单额");
		cell19456 = row19456.createCell((short) 5);
		cell19456.setCellValue("美式");
		cell19456 = row19456.createCell((short) 6);
		cell19456.setCellValue("拿铁");
		cell19456 = row19456.createCell((short) 7);
		cell19456.setCellValue("摩卡");
		cell19456 = row19456.createCell((short) 8);
		cell19456.setCellValue("卡布奇诺");
		cell19456 = row19456.createCell((short) 9);
		cell19456.setCellValue("巧克力");
		cell19456 = row19456.createCell((short) 10);
		cell19456.setCellValue("玛琪雅朵");
		cell19456 = row19456.createCell((short) 11);
		cell19456.setCellValue("牛奶");
		cell19456 = row19456.createCell((short) 12);
		cell19456.setCellValue("巧克力牛奶");
		cell19456 = row19456.createCell((short) 13);
		cell19456.setCellValue("抹茶");
		cell19456 = row19456.createCell((short) 14);
		cell19456.setCellValue("抹茶咖啡");

		HSSFRow row19457 = sheet19457.createRow((int) 0);
		HSSFCell cell19457 = row19457.createCell((short) 0);
		cell19457.setCellValue("日期");
		cell19457 = row19457.createCell((short) 1);
		cell19457.setCellValue("付费杯数");
		cell19457 = row19457.createCell((short) 2);
		cell19457.setCellValue("收费金额");
		cell19457 = row19457.createCell((short) 3);
		cell19457.setCellValue("免费杯数");
		cell19457 = row19457.createCell((short) 4);
		cell19457.setCellValue("免单额");
		cell19457 = row19457.createCell((short) 5);
		cell19457.setCellValue("美式");
		cell19457 = row19457.createCell((short) 6);
		cell19457.setCellValue("拿铁");
		cell19457 = row19457.createCell((short) 7);
		cell19457.setCellValue("摩卡");
		cell19457 = row19457.createCell((short) 8);
		cell19457.setCellValue("卡布奇诺");
		cell19457 = row19457.createCell((short) 9);
		cell19457.setCellValue("巧克力");
		cell19457 = row19457.createCell((short) 10);
		cell19457.setCellValue("玛琪雅朵");
		cell19457 = row19457.createCell((short) 11);
		cell19457.setCellValue("牛奶");
		cell19457 = row19457.createCell((short) 12);
		cell19457.setCellValue("巧克力牛奶");
		cell19457 = row19457.createCell((short) 13);
		cell19457.setCellValue("抹茶");
		cell19457 = row19457.createCell((short) 14);
		cell19457.setCellValue("抹茶咖啡");

		HSSFRow row19458 = sheet19458.createRow((int) 0);
		HSSFCell cell19458 = row19458.createCell((short) 0);
		cell19458.setCellValue("日期");
		cell19458 = row19458.createCell((short) 1);
		cell19458.setCellValue("付费杯数");
		cell19458 = row19458.createCell((short) 2);
		cell19458.setCellValue("收费金额");
		cell19458 = row19458.createCell((short) 3);
		cell19458.setCellValue("免费杯数");
		cell19458 = row19458.createCell((short) 4);
		cell19458.setCellValue("免单额");
		cell19458 = row19458.createCell((short) 5);
		cell19458.setCellValue("美式");
		cell19458 = row19458.createCell((short) 6);
		cell19458.setCellValue("拿铁");
		cell19458 = row19458.createCell((short) 7);
		cell19458.setCellValue("摩卡");
		cell19458 = row19458.createCell((short) 8);
		cell19458.setCellValue("卡布奇诺");
		cell19458 = row19458.createCell((short) 9);
		cell19458.setCellValue("巧克力");
		cell19458 = row19458.createCell((short) 10);
		cell19458.setCellValue("玛琪雅朵");
		cell19458 = row19458.createCell((short) 11);
		cell19458.setCellValue("牛奶");
		cell19458 = row19458.createCell((short) 12);
		cell19458.setCellValue("巧克力牛奶");
		cell19458 = row19458.createCell((short) 13);
		cell19458.setCellValue("抹茶");
		cell19458 = row19458.createCell((short) 14);
		cell19458.setCellValue("抹茶咖啡");

		HSSFRow row19459 = sheet19459.createRow((int) 0);
		HSSFCell cell19459 = row19459.createCell((short) 0);
		cell19459.setCellValue("日期");
		cell19459 = row19459.createCell((short) 1);
		cell19459.setCellValue("付费杯数");
		cell19459 = row19459.createCell((short) 2);
		cell19459.setCellValue("收费金额");
		cell19459 = row19459.createCell((short) 3);
		cell19459.setCellValue("免费杯数");
		cell19459 = row19459.createCell((short) 4);
		cell19459.setCellValue("免单额");
		cell19459 = row19459.createCell((short) 5);
		cell19459.setCellValue("美式");
		cell19459 = row19459.createCell((short) 6);
		cell19459.setCellValue("拿铁");
		cell19459 = row19459.createCell((short) 7);
		cell19459.setCellValue("摩卡");
		cell19459 = row19459.createCell((short) 8);
		cell19459.setCellValue("卡布奇诺");
		cell19459 = row19459.createCell((short) 9);
		cell19459.setCellValue("巧克力");
		cell19459 = row19459.createCell((short) 10);
		cell19459.setCellValue("玛琪雅朵");
		cell19459 = row19459.createCell((short) 11);
		cell19459.setCellValue("牛奶");
		cell19459 = row19459.createCell((short) 12);
		cell19459.setCellValue("巧克力牛奶");
		cell19459 = row19459.createCell((short) 13);
		cell19459.setCellValue("抹茶");
		cell19459 = row19459.createCell((short) 14);
		cell19459.setCellValue("抹茶咖啡");

		HSSFRow row19472 = sheet19472.createRow((int) 0);
		HSSFCell cell19472 = row19472.createCell((short) 0);
		cell19472.setCellValue("日期");
		cell19472 = row19472.createCell((short) 1);
		cell19472.setCellValue("付费杯数");
		cell19472 = row19472.createCell((short) 2);
		cell19472.setCellValue("收费金额");
		cell19472 = row19472.createCell((short) 3);
		cell19472.setCellValue("免费杯数");
		cell19472 = row19472.createCell((short) 4);
		cell19472.setCellValue("免单额");
		cell19472 = row19472.createCell((short) 5);
		cell19472.setCellValue("美式");
		cell19472 = row19472.createCell((short) 6);
		cell19472.setCellValue("拿铁");
		cell19472 = row19472.createCell((short) 7);
		cell19472.setCellValue("摩卡");
		cell19472 = row19472.createCell((short) 8);
		cell19472.setCellValue("卡布奇诺");
		cell19472 = row19472.createCell((short) 9);
		cell19472.setCellValue("巧克力");
		cell19472 = row19472.createCell((short) 10);
		cell19472.setCellValue("玛琪雅朵");
		cell19472 = row19472.createCell((short) 11);
		cell19472.setCellValue("牛奶");
		cell19472 = row19472.createCell((short) 12);
		cell19472.setCellValue("巧克力牛奶");
		cell19472 = row19472.createCell((short) 13);
		cell19472.setCellValue("抹茶");
		cell19472 = row19472.createCell((short) 14);
		cell19472.setCellValue("抹茶咖啡");

		HSSFRow row19473 = sheet19473.createRow((int) 0);
		HSSFCell cell19473 = row19473.createCell((short) 0);
		cell19473.setCellValue("日期");
		cell19473 = row19473.createCell((short) 1);
		cell19473.setCellValue("付费杯数");
		cell19473 = row19473.createCell((short) 2);
		cell19473.setCellValue("收费金额");
		cell19473 = row19473.createCell((short) 3);
		cell19473.setCellValue("免费杯数");
		cell19473 = row19473.createCell((short) 4);
		cell19473.setCellValue("免单额");
		cell19473 = row19473.createCell((short) 5);
		cell19473.setCellValue("美式");
		cell19473 = row19473.createCell((short) 6);
		cell19473.setCellValue("拿铁");
		cell19473 = row19473.createCell((short) 7);
		cell19473.setCellValue("摩卡");
		cell19473 = row19473.createCell((short) 8);
		cell19473.setCellValue("卡布奇诺");
		cell19473 = row19473.createCell((short) 9);
		cell19473.setCellValue("巧克力");
		cell19473 = row19473.createCell((short) 10);
		cell19473.setCellValue("玛琪雅朵");
		cell19473 = row19473.createCell((short) 11);
		cell19473.setCellValue("牛奶");
		cell19473 = row19473.createCell((short) 12);
		cell19473.setCellValue("巧克力牛奶");
		cell19473 = row19473.createCell((short) 13);
		cell19473.setCellValue("抹茶");
		cell19473 = row19473.createCell((short) 14);
		cell19473.setCellValue("抹茶咖啡");

		HSSFRow row19475 = sheet19475.createRow((int) 0);
		HSSFCell cell19475 = row19475.createCell((short) 0);
		cell19475.setCellValue("日期");
		cell19475 = row19475.createCell((short) 1);
		cell19475.setCellValue("付费杯数");
		cell19475 = row19475.createCell((short) 2);
		cell19475.setCellValue("收费金额");
		cell19475 = row19475.createCell((short) 3);
		cell19475.setCellValue("免费杯数");
		cell19475 = row19475.createCell((short) 4);
		cell19475.setCellValue("免单额");
		cell19475 = row19475.createCell((short) 5);
		cell19475.setCellValue("美式");
		cell19475 = row19475.createCell((short) 6);
		cell19475.setCellValue("拿铁");
		cell19475 = row19475.createCell((short) 7);
		cell19475.setCellValue("摩卡");
		cell19475 = row19475.createCell((short) 8);
		cell19475.setCellValue("卡布奇诺");
		cell19475 = row19475.createCell((short) 9);
		cell19475.setCellValue("巧克力");
		cell19475 = row19475.createCell((short) 10);
		cell19475.setCellValue("玛琪雅朵");
		cell19475 = row19475.createCell((short) 11);
		cell19475.setCellValue("牛奶");
		cell19475 = row19475.createCell((short) 12);
		cell19475.setCellValue("巧克力牛奶");
		cell19475 = row19475.createCell((short) 13);
		cell19475.setCellValue("抹茶");
		cell19475 = row19475.createCell((short) 14);
		cell19475.setCellValue("抹茶咖啡");

		HSSFRow row19481 = sheet19481.createRow((int) 0);
		HSSFCell cell19481 = row19481.createCell((short) 0);
		cell19481.setCellValue("日期");
		cell19481 = row19481.createCell((short) 1);
		cell19481.setCellValue("付费杯数");
		cell19481 = row19481.createCell((short) 2);
		cell19481.setCellValue("收费金额");
		cell19481 = row19481.createCell((short) 3);
		cell19481.setCellValue("免费杯数");
		cell19481 = row19481.createCell((short) 4);
		cell19481.setCellValue("免单额");
		cell19481 = row19481.createCell((short) 5);
		cell19481.setCellValue("美式");
		cell19481 = row19481.createCell((short) 6);
		cell19481.setCellValue("拿铁");
		cell19481 = row19481.createCell((short) 7);
		cell19481.setCellValue("摩卡");
		cell19481 = row19481.createCell((short) 8);
		cell19481.setCellValue("卡布奇诺");
		cell19481 = row19481.createCell((short) 9);
		cell19481.setCellValue("巧克力");
		cell19481 = row19481.createCell((short) 10);
		cell19481.setCellValue("玛琪雅朵");
		cell19481 = row19481.createCell((short) 11);
		cell19481.setCellValue("牛奶");
		cell19481 = row19481.createCell((short) 12);
		cell19481.setCellValue("巧克力牛奶");
		cell19481 = row19481.createCell((short) 13);
		cell19481.setCellValue("抹茶");
		cell19481 = row19481.createCell((short) 14);
		cell19481.setCellValue("抹茶咖啡");

		HSSFRow row19482 = sheet19482.createRow((int) 0);
		HSSFCell cell19482 = row19482.createCell((short) 0);
		cell19482.setCellValue("日期");
		cell19482 = row19482.createCell((short) 1);
		cell19482.setCellValue("付费杯数");
		cell19482 = row19482.createCell((short) 2);
		cell19482.setCellValue("收费金额");
		cell19482 = row19482.createCell((short) 3);
		cell19482.setCellValue("免费杯数");
		cell19482 = row19482.createCell((short) 4);
		cell19482.setCellValue("免单额");
		cell19482 = row19482.createCell((short) 5);
		cell19482.setCellValue("美式");
		cell19482 = row19482.createCell((short) 6);
		cell19482.setCellValue("拿铁");
		cell19482 = row19482.createCell((short) 7);
		cell19482.setCellValue("摩卡");
		cell19482 = row19482.createCell((short) 8);
		cell19482.setCellValue("卡布奇诺");
		cell19482 = row19482.createCell((short) 9);
		cell19482.setCellValue("巧克力");
		cell19482 = row19482.createCell((short) 10);
		cell19482.setCellValue("玛琪雅朵");
		cell19482 = row19482.createCell((short) 11);
		cell19482.setCellValue("牛奶");
		cell19482 = row19482.createCell((short) 12);
		cell19482.setCellValue("巧克力牛奶");
		cell19482 = row19482.createCell((short) 13);
		cell19482.setCellValue("抹茶");
		cell19482 = row19482.createCell((short) 14);
		cell19482.setCellValue("抹茶咖啡");

		HSSFRow row19483 = sheet19483.createRow((int) 0);
		HSSFCell cell19483 = row19483.createCell((short) 0);
		cell19483.setCellValue("日期");
		cell19483 = row19483.createCell((short) 1);
		cell19483.setCellValue("付费杯数");
		cell19483 = row19483.createCell((short) 2);
		cell19483.setCellValue("收费金额");
		cell19483 = row19483.createCell((short) 3);
		cell19483.setCellValue("免费杯数");
		cell19483 = row19483.createCell((short) 4);
		cell19483.setCellValue("免单额");
		cell19483 = row19483.createCell((short) 5);
		cell19483.setCellValue("美式");
		cell19483 = row19483.createCell((short) 6);
		cell19483.setCellValue("拿铁");
		cell19483 = row19483.createCell((short) 7);
		cell19483.setCellValue("摩卡");
		cell19483 = row19483.createCell((short) 8);
		cell19483.setCellValue("卡布奇诺");
		cell19483 = row19483.createCell((short) 9);
		cell19483.setCellValue("巧克力");
		cell19483 = row19483.createCell((short) 10);
		cell19483.setCellValue("玛琪雅朵");
		cell19483 = row19483.createCell((short) 11);
		cell19483.setCellValue("牛奶");
		cell19483 = row19483.createCell((short) 12);
		cell19483.setCellValue("巧克力牛奶");
		cell19483 = row19483.createCell((short) 13);
		cell19483.setCellValue("抹茶");
		cell19483 = row19483.createCell((short) 14);
		cell19483.setCellValue("抹茶咖啡");

		HSSFRow row19484 = sheet19484.createRow((int) 0);
		HSSFCell cell19484 = row19484.createCell((short) 0);
		cell19484.setCellValue("日期");
		cell19484 = row19484.createCell((short) 1);
		cell19484.setCellValue("付费杯数");
		cell19484 = row19484.createCell((short) 2);
		cell19484.setCellValue("收费金额");
		cell19484 = row19484.createCell((short) 3);
		cell19484.setCellValue("免费杯数");
		cell19484 = row19484.createCell((short) 4);
		cell19484.setCellValue("免单额");
		cell19484 = row19484.createCell((short) 5);
		cell19484.setCellValue("美式");
		cell19484 = row19484.createCell((short) 6);
		cell19484.setCellValue("拿铁");
		cell19484 = row19484.createCell((short) 7);
		cell19484.setCellValue("摩卡");
		cell19484 = row19484.createCell((short) 8);
		cell19484.setCellValue("卡布奇诺");
		cell19484 = row19484.createCell((short) 9);
		cell19484.setCellValue("巧克力");
		cell19484 = row19484.createCell((short) 10);
		cell19484.setCellValue("玛琪雅朵");
		cell19484 = row19484.createCell((short) 11);
		cell19484.setCellValue("牛奶");
		cell19484 = row19484.createCell((short) 12);
		cell19484.setCellValue("巧克力牛奶");
		cell19484 = row19484.createCell((short) 13);
		cell19484.setCellValue("抹茶");
		cell19484 = row19484.createCell((short) 14);
		cell19484.setCellValue("抹茶咖啡");

		HSSFRow rowDaySum = sheetSum.createRow((int) 0);
		HSSFCell cellDaySum = rowDaySum.createCell((short) 0);
		cellDaySum.setCellValue("设备名称");
		cellDaySum = rowDaySum.createCell((short) 1);
		cellDaySum.setCellValue("付费杯数");
		cellDaySum = rowDaySum.createCell((short) 2);
		cellDaySum.setCellValue("收费金额");
		cellDaySum = rowDaySum.createCell((short) 3);
		cellDaySum.setCellValue("免费杯数");
		cellDaySum = rowDaySum.createCell((short) 4);
		cellDaySum.setCellValue("免单额");
		cellDaySum = rowDaySum.createCell((short) 5);
		cellDaySum.setCellValue("美式");
		cellDaySum = rowDaySum.createCell((short) 6);
		cellDaySum.setCellValue("拿铁");
		cellDaySum = rowDaySum.createCell((short) 7);
		cellDaySum.setCellValue("摩卡");
		cellDaySum = rowDaySum.createCell((short) 8);
		cellDaySum.setCellValue("卡布奇诺");
		cellDaySum = rowDaySum.createCell((short) 9);
		cellDaySum.setCellValue("巧克力");
		cellDaySum = rowDaySum.createCell((short) 10);
		cellDaySum.setCellValue("玛琪雅朵");
		cellDaySum = rowDaySum.createCell((short) 11);
		cellDaySum.setCellValue("牛奶");
		cellDaySum = rowDaySum.createCell((short) 12);
		cellDaySum.setCellValue("巧克力牛奶");
		cellDaySum = rowDaySum.createCell((short) 13);
		cellDaySum.setCellValue("抹茶");
		cellDaySum = rowDaySum.createCell((short) 14);
		cellDaySum.setCellValue("抹茶咖啡");

		int payNum18640 = 0;
		double payMoney18640 = 0;
		int freeNum18640 = 0;
		int meiShi18640 = 0;
		int naTie18640 = 0;
		int moKa18640 = 0;
		int KaBu18640 = 0;
		int qiaoKe18640 = 0;
		int maQi18640 = 0;
		int niuNa18640 = 0;
		int qiaoKeNai18640 = 0;
		int moCha18640 = 0;
		int moChaKa18640 = 0;

		int payNum19452 = 0;
		double payMoney19452 = 0;
		int freeNum19452 = 0;
		int meiShi19452 = 0;
		int naTie19452 = 0;
		int moKa19452 = 0;
		int KaBu19452 = 0;
		int qiaoKe19452 = 0;
		int maQi19452 = 0;
		int niuNa19452 = 0;
		int qiaoKeNai19452 = 0;
		int moCha19452 = 0;
		int moChaKa19452 = 0;

		int payNum19453 = 0;
		double payMoney19453 = 0;
		int freeNum19453 = 0;
		int meiShi19453 = 0;
		int naTie19453 = 0;
		int moKa19453 = 0;
		int KaBu19453 = 0;
		int qiaoKe19453 = 0;
		int maQi19453 = 0;
		int niuNa19453 = 0;
		int qiaoKeNai19453 = 0;
		int moCha19453 = 0;
		int moChaKa19453 = 0;

		int payNum19454 = 0;
		double payMoney19454 = 0;
		int freeNum19454 = 0;
		int meiShi19454 = 0;
		int naTie19454 = 0;
		int moKa19454 = 0;
		int KaBu19454 = 0;
		int qiaoKe19454 = 0;
		int maQi19454 = 0;
		int niuNa19454 = 0;
		int qiaoKeNai19454 = 0;
		int moCha19454 = 0;
		int moChaKa19454 = 0;

		int payNum19455 = 0;
		double payMoney19455 = 0;
		int freeNum19455 = 0;
		int meiShi19455 = 0;
		int naTie19455 = 0;
		int moKa19455 = 0;
		int KaBu19455 = 0;
		int qiaoKe19455 = 0;
		int maQi19455 = 0;
		int niuNa19455 = 0;
		int qiaoKeNai19455 = 0;
		int moCha19455 = 0;
		int moChaKa19455 = 0;

		int payNum19456 = 0;
		double payMoney19456 = 0;
		int freeNum19456 = 0;
		int meiShi19456 = 0;
		int naTie19456 = 0;
		int moKa19456 = 0;
		int KaBu19456 = 0;
		int qiaoKe19456 = 0;
		int maQi19456 = 0;
		int niuNa19456 = 0;
		int qiaoKeNai19456 = 0;
		int moCha19456 = 0;
		int moChaKa19456 = 0;

		int payNum19457 = 0;
		double payMoney19457 = 0;
		int freeNum19457 = 0;
		int meiShi19457 = 0;
		int naTie19457 = 0;
		int moKa19457 = 0;
		int KaBu19457 = 0;
		int qiaoKe19457 = 0;
		int maQi19457 = 0;
		int niuNa19457 = 0;
		int qiaoKeNai19457 = 0;
		int moCha19457 = 0;
		int moChaKa19457 = 0;

		int payNum19458 = 0;
		double payMoney19458 = 0;
		int freeNum19458 = 0;
		int meiShi19458 = 0;
		int naTie19458 = 0;
		int moKa19458 = 0;
		int KaBu19458 = 0;
		int qiaoKe19458 = 0;
		int maQi19458 = 0;
		int niuNa19458 = 0;
		int qiaoKeNai19458 = 0;
		int moCha19458 = 0;
		int moChaKa19458 = 0;

		int payNum19459 = 0;
		double payMoney19459 = 0;
		int freeNum19459 = 0;
		int meiShi19459 = 0;
		int naTie19459 = 0;
		int moKa19459 = 0;
		int KaBu19459 = 0;
		int qiaoKe19459 = 0;
		int maQi19459 = 0;
		int niuNa19459 = 0;
		int qiaoKeNai19459 = 0;
		int moCha19459 = 0;
		int moChaKa19459 = 0;

		int payNum19472 = 0;
		double payMoney19472 = 0;
		int freeNum19472 = 0;
		int meiShi19472 = 0;
		int naTie19472 = 0;
		int moKa19472 = 0;
		int KaBu19472 = 0;
		int qiaoKe19472 = 0;
		int maQi19472 = 0;
		int niuNa19472 = 0;
		int qiaoKeNai19472 = 0;
		int moCha19472 = 0;
		int moChaKa19472 = 0;

		int payNum19473 = 0;
		double payMoney19473 = 0;
		int freeNum19473 = 0;
		int meiShi19473 = 0;
		int naTie19473 = 0;
		int moKa19473 = 0;
		int KaBu19473 = 0;
		int qiaoKe19473 = 0;
		int maQi19473 = 0;
		int niuNa19473 = 0;
		int qiaoKeNai19473 = 0;
		int moCha19473 = 0;
		int moChaKa19473 = 0;

		int payNum19475 = 0;
		double payMoney19475 = 0;
		int freeNum19475 = 0;
		int meiShi19475 = 0;
		int naTie19475 = 0;
		int moKa19475 = 0;
		int KaBu19475 = 0;
		int qiaoKe19475 = 0;
		int maQi19475 = 0;
		int niuNa19475 = 0;
		int qiaoKeNai19475 = 0;
		int moCha19475 = 0;
		int moChaKa19475 = 0;

		int payNum19481 = 0;
		double payMoney19481 = 0;
		int freeNum19481 = 0;
		int meiShi19481 = 0;
		int naTie19481 = 0;
		int moKa19481 = 0;
		int KaBu19481 = 0;
		int qiaoKe19481 = 0;
		int maQi19481 = 0;
		int niuNa19481 = 0;
		int qiaoKeNai19481 = 0;
		int moCha19481 = 0;
		int moChaKa19481 = 0;

		int payNum19482 = 0;
		double payMoney19482 = 0;
		int freeNum19482 = 0;
		int meiShi19482 = 0;
		int naTie19482 = 0;
		int moKa19482 = 0;
		int KaBu19482 = 0;
		int qiaoKe19482 = 0;
		int maQi19482 = 0;
		int niuNa19482 = 0;
		int qiaoKeNai19482 = 0;
		int moCha19482 = 0;
		int moChaKa19482 = 0;

		int payNum19483 = 0;
		double payMoney19483 = 0;
		int freeNum19483 = 0;
		int meiShi19483 = 0;
		int naTie19483 = 0;
		int moKa19483 = 0;
		int KaBu19483 = 0;
		int qiaoKe19483 = 0;
		int maQi19483 = 0;
		int niuNa19483 = 0;
		int qiaoKeNai19483 = 0;
		int moCha19483 = 0;
		int moChaKa19483 = 0;

		int payNum19484 = 0;
		double payMoney19484 = 0;
		int freeNum19484 = 0;
		int meiShi19484 = 0;
		int naTie19484 = 0;
		int moKa19484 = 0;
		int KaBu19484 = 0;
		int qiaoKe19484 = 0;
		int maQi19484 = 0;
		int niuNa19484 = 0;
		int qiaoKeNai19484 = 0;
		int moCha19484 = 0;
		int moChaKa19484 = 0;

		int payNumTotal = 0;
		double payMoneyTotal = 0;
		int freeNumTotal = 0;
		int meiShiTotal = 0;
		int naTieTotal = 0;
		int moKaTotal = 0;
		int KaBuTotal = 0;
		int qiaoKeTotal = 0;
		int maQiTotal = 0;
		int niuNaTotal = 0;
		int qiaoKeNaiTotal = 0;
		int moChaTotal = 0;
		int moChaKaTotal = 0;

		for (Map.Entry<String, List<Infos>> entry : fileData.entrySet()) {

			List<Infos> infos = (List<Infos>) entry.getValue();
			int lineNum = infos.size();

			for (int i = 0; i < infos.size(); i++) {

				if (entry.getKey().equals("18640")) {

					row18640 = sheet18640.createRow((int) i + 1);
					Infos info = infos.get(i);
					row18640.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row18640.createCell((short) 1).setCellValue(1);
						row18640.createCell((short) 2).setCellValue(info.getpMoney());
						payNum18640++;
						payMoney18640 = payMoney18640 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();

					} else {
						row18640.createCell((short) 3).setCellValue(1);
						freeNum18640++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row18640.createCell((short) 5).setCellValue(1);
						meiShi18640++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row18640.createCell((short) 6).setCellValue(1);
						naTie18640++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row18640.createCell((short) 7).setCellValue(1);
						moKa18640++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row18640.createCell((short) 8).setCellValue(1);
						KaBu18640++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row18640.createCell((short) 9).setCellValue(1);
						qiaoKe18640++;
						qiaoKeTotal++;

					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row18640.createCell((short) 10).setCellValue(1);
						maQi18640++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row18640.createCell((short) 11).setCellValue(1);
						niuNa18640++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row18640.createCell((short) 12).setCellValue(1);
						qiaoKeNai18640++;
						qiaoKeNaiTotal++;

					}
					if (info.getpName().equals("抹茶")) {
						row18640.createCell((short) 13).setCellValue(1);
						moCha18640++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row18640.createCell((short) 14).setCellValue(1);
						moChaKa18640++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19452")) {
					row19452 = sheet19452.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19452.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19452.createCell((short) 1).setCellValue(1);
						row19452.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19452++;
						payMoney19452 = payMoney19452 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19452.createCell((short) 3).setCellValue(1);
						freeNum19452++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19452.createCell((short) 5).setCellValue(1);
						meiShi19452++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19452.createCell((short) 6).setCellValue(1);
						naTie19452++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19452.createCell((short) 7).setCellValue(1);
						moKa19452++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19452.createCell((short) 8).setCellValue(1);
						KaBu19452++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19452.createCell((short) 9).setCellValue(1);
						qiaoKe19452++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19452.createCell((short) 10).setCellValue(1);
						maQi19452++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19452.createCell((short) 11).setCellValue(1);
						niuNa19452++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19452.createCell((short) 12).setCellValue(1);
						qiaoKeNai19452++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19452.createCell((short) 13).setCellValue(1);
						moCha19452++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19452.createCell((short) 14).setCellValue(1);
						moChaKa19452++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19453")) {
					row19453 = sheet19453.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19453.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19453.createCell((short) 1).setCellValue(1);
						row19453.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19453++;
						payMoney19453 = payMoney19453 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19453.createCell((short) 3).setCellValue(1);
						freeNum19453++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19453.createCell((short) 5).setCellValue(1);
						meiShi19453++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19453.createCell((short) 6).setCellValue(1);
						naTie19453++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19453.createCell((short) 7).setCellValue(1);
						moKa19453++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19453.createCell((short) 8).setCellValue(1);
						KaBu19453++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19453.createCell((short) 9).setCellValue(1);
						qiaoKe19453++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19453.createCell((short) 10).setCellValue(1);
						maQi19453++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19453.createCell((short) 11).setCellValue(1);
						niuNa19453++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19453.createCell((short) 12).setCellValue(1);
						qiaoKeNai19453++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19453.createCell((short) 13).setCellValue(1);
						moCha19453++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19453.createCell((short) 14).setCellValue(1);
						moChaKa19453++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19454")) {
					row19454 = sheet19454.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19454.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19454.createCell((short) 1).setCellValue(1);
						row19454.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19454++;
						payMoney19454 = payMoney19454 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19454.createCell((short) 3).setCellValue(1);
						freeNum19454++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19454.createCell((short) 5).setCellValue(1);
						meiShi19454++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19454.createCell((short) 6).setCellValue(1);
						naTie19454++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19454.createCell((short) 7).setCellValue(1);
						moKa19454++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19454.createCell((short) 8).setCellValue(1);
						KaBu19454++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19454.createCell((short) 9).setCellValue(1);
						qiaoKe19454++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19454.createCell((short) 10).setCellValue(1);
						maQi19454++;
						maQiTotal++;

					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19454.createCell((short) 11).setCellValue(1);
						niuNa19454++;
						niuNaTotal++;

					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19454.createCell((short) 12).setCellValue(1);
						qiaoKeNai19454++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19454.createCell((short) 13).setCellValue(1);
						moCha19454++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19454.createCell((short) 14).setCellValue(1);
						moChaKa19454++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19455")) {
					row19455 = sheet19455.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19455.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19455.createCell((short) 1).setCellValue(1);
						row19455.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19455++;
						payMoney19455 = payMoney19455 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19455.createCell((short) 3).setCellValue(1);
						freeNum19455++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19455.createCell((short) 5).setCellValue(1);
						meiShi19455++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19455.createCell((short) 6).setCellValue(1);
						naTie19455++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19455.createCell((short) 7).setCellValue(1);
						moKa19455++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19455.createCell((short) 8).setCellValue(1);
						KaBu19455++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19455.createCell((short) 9).setCellValue(1);
						qiaoKe19455++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19455.createCell((short) 10).setCellValue(1);
						maQi19455++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19455.createCell((short) 11).setCellValue(1);
						niuNa19455++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19455.createCell((short) 12).setCellValue(1);
						qiaoKeNai19455++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19455.createCell((short) 13).setCellValue(1);
						moCha19455++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19455.createCell((short) 14).setCellValue(1);
						moChaKa19455++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19456")) {
					row19456 = sheet19456.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19456.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19456.createCell((short) 1).setCellValue(1);
						row19456.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19456++;
						payMoney19456 = payMoney19456 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19456.createCell((short) 3).setCellValue(1);
						freeNum19456++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19456.createCell((short) 5).setCellValue(1);
						meiShi19456++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19456.createCell((short) 6).setCellValue(1);
						naTie19456++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19456.createCell((short) 7).setCellValue(1);
						moKa19456++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19456.createCell((short) 8).setCellValue(1);
						KaBu19456++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19456.createCell((short) 9).setCellValue(1);
						qiaoKe19456++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19456.createCell((short) 10).setCellValue(1);
						maQi19456++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19456.createCell((short) 11).setCellValue(1);
						niuNa19456++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19456.createCell((short) 12).setCellValue(1);
						qiaoKeNai19456++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19456.createCell((short) 13).setCellValue(1);
						moCha19456++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19456.createCell((short) 14).setCellValue(1);
						moChaKa19456++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19457")) {
					row19457 = sheet19457.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19457.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19457.createCell((short) 1).setCellValue(1);
						row19457.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19457++;
						payMoney19457 = payMoney19457 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19457.createCell((short) 3).setCellValue(1);
						freeNum19457++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19457.createCell((short) 5).setCellValue(1);
						meiShi19457++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19457.createCell((short) 6).setCellValue(1);
						naTie19457++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19457.createCell((short) 7).setCellValue(1);
						moKa19457++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19457.createCell((short) 8).setCellValue(1);
						KaBu19457++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19457.createCell((short) 9).setCellValue(1);
						qiaoKe19457++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19457.createCell((short) 10).setCellValue(1);
						maQi19457++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19457.createCell((short) 11).setCellValue(1);
						niuNa19457++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19457.createCell((short) 12).setCellValue(1);
						qiaoKeNai19457++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19457.createCell((short) 13).setCellValue(1);
						moCha19457++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19457.createCell((short) 14).setCellValue(1);
						moChaKa19457++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19458")) {
					row19458 = sheet19458.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19458.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19458.createCell((short) 1).setCellValue(1);
						row19458.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19458++;
						payMoney19458 = payMoney19458 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19458.createCell((short) 3).setCellValue(1);
						freeNum19458++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19458.createCell((short) 5).setCellValue(1);
						meiShi19458++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19458.createCell((short) 6).setCellValue(1);
						naTie19458++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19458.createCell((short) 7).setCellValue(1);
						moKa19458++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19458.createCell((short) 8).setCellValue(1);
						KaBu19458++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19458.createCell((short) 9).setCellValue(1);
						qiaoKe19458++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19458.createCell((short) 10).setCellValue(1);
						maQi19458++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19458.createCell((short) 11).setCellValue(1);
						niuNa19458++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19458.createCell((short) 12).setCellValue(1);
						qiaoKeNai19458++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19458.createCell((short) 13).setCellValue(1);
						moCha19458++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19458.createCell((short) 14).setCellValue(1);
						moChaKa19458++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19459")) {
					row19459 = sheet19459.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19459.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19459.createCell((short) 1).setCellValue(1);
						row19459.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19459++;
						payMoney19459 = payMoney19459 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19459.createCell((short) 3).setCellValue(1);
						freeNum19459++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19459.createCell((short) 5).setCellValue(1);
						meiShi19459++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19459.createCell((short) 6).setCellValue(1);
						naTie19459++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19459.createCell((short) 7).setCellValue(1);
						moKa19459++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19459.createCell((short) 8).setCellValue(1);
						KaBu19459++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19459.createCell((short) 9).setCellValue(1);
						qiaoKe19459++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19459.createCell((short) 10).setCellValue(1);
						maQi19459++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19459.createCell((short) 11).setCellValue(1);
						niuNa19459++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19459.createCell((short) 12).setCellValue(1);
						qiaoKeNai19459++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19459.createCell((short) 13).setCellValue(1);
						moCha19459++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19459.createCell((short) 14).setCellValue(1);
						moChaKa19459++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19472")) {
					row19472 = sheet19472.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19472.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19472.createCell((short) 1).setCellValue(1);
						row19472.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19472++;
						payMoney19472 = payMoney19472 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19472.createCell((short) 3).setCellValue(1);
						freeNum19472++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19472.createCell((short) 5).setCellValue(1);
						meiShi19472++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19472.createCell((short) 6).setCellValue(1);
						naTie19472++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19472.createCell((short) 7).setCellValue(1);
						moKa19472++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19472.createCell((short) 8).setCellValue(1);
						KaBu19472++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19472.createCell((short) 9).setCellValue(1);
						qiaoKe19472++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19472.createCell((short) 10).setCellValue(1);
						maQi19472++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19472.createCell((short) 11).setCellValue(1);
						niuNa19472++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19472.createCell((short) 12).setCellValue(1);
						qiaoKeNai19472++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19472.createCell((short) 13).setCellValue(1);
						moCha19472++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19472.createCell((short) 14).setCellValue(1);
						moChaKa19472++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19473")) {
					row19473 = sheet19473.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19473.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19473.createCell((short) 1).setCellValue(1);
						row19473.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19473++;
						payMoney19473 = payMoney19473 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19473.createCell((short) 3).setCellValue(1);
						freeNum19473++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19473.createCell((short) 5).setCellValue(1);
						meiShi19473++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19473.createCell((short) 6).setCellValue(1);
						naTie19473++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19473.createCell((short) 7).setCellValue(1);
						moKa19473++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19473.createCell((short) 8).setCellValue(1);
						KaBu19473++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19473.createCell((short) 9).setCellValue(1);
						qiaoKe19473++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19473.createCell((short) 10).setCellValue(1);
						maQi19473++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19473.createCell((short) 11).setCellValue(1);
						niuNa19473++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19473.createCell((short) 12).setCellValue(1);
						qiaoKeNai19473++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19473.createCell((short) 13).setCellValue(1);
						moCha19473++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19473.createCell((short) 14).setCellValue(1);
						moChaKa19473++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19475")) {
					row19475 = sheet19475.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19475.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19475.createCell((short) 1).setCellValue(1);
						row19475.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19475++;
						payMoney19475 = payMoney19475 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19475.createCell((short) 3).setCellValue(1);
						freeNum19475++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19475.createCell((short) 5).setCellValue(1);
						meiShi19475++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19475.createCell((short) 6).setCellValue(1);
						naTie19475++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19475.createCell((short) 7).setCellValue(1);
						moKa19475++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19475.createCell((short) 8).setCellValue(1);
						KaBu19475++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19475.createCell((short) 9).setCellValue(1);
						qiaoKe19475++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19475.createCell((short) 10).setCellValue(1);
						maQi19475++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19475.createCell((short) 11).setCellValue(1);
						niuNa19475++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19475.createCell((short) 12).setCellValue(1);
						qiaoKeNai19475++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19475.createCell((short) 13).setCellValue(1);
						moCha19475++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19475.createCell((short) 14).setCellValue(1);
						moChaKa19475++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19481")) {
					row19481 = sheet19481.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19481.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19481.createCell((short) 1).setCellValue(1);
						row19481.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19481++;
						payMoney19481 = payMoney19481 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19481.createCell((short) 3).setCellValue(1);
						freeNum19481++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19481.createCell((short) 5).setCellValue(1);
						meiShi19481++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19481.createCell((short) 6).setCellValue(1);
						naTie19481++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19481.createCell((short) 7).setCellValue(1);
						moKa19481++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19481.createCell((short) 8).setCellValue(1);
						KaBu19481++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19481.createCell((short) 9).setCellValue(1);
						qiaoKe19481++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19481.createCell((short) 10).setCellValue(1);
						maQi19481++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19481.createCell((short) 11).setCellValue(1);
						niuNa19481++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19481.createCell((short) 12).setCellValue(1);
						qiaoKeNai19481++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19481.createCell((short) 13).setCellValue(1);
						moCha19481++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19481.createCell((short) 14).setCellValue(1);
						moChaKa19481++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19482")) {
					row19482 = sheet19482.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19482.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19482.createCell((short) 1).setCellValue(1);
						row19482.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19482++;
						payMoney19482 = payMoney19482 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19482.createCell((short) 3).setCellValue(1);
						freeNum19482++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19482.createCell((short) 5).setCellValue(1);
						meiShi19482++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19482.createCell((short) 6).setCellValue(1);
						naTie19482++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19482.createCell((short) 7).setCellValue(1);
						moKa19482++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19482.createCell((short) 8).setCellValue(1);
						KaBu19482++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19482.createCell((short) 9).setCellValue(1);
						qiaoKe19482++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19482.createCell((short) 10).setCellValue(1);
						maQi19482++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19482.createCell((short) 11).setCellValue(1);
						niuNa19482++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19482.createCell((short) 12).setCellValue(1);
						qiaoKeNai19482++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19482.createCell((short) 13).setCellValue(1);
						moCha19482++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19482.createCell((short) 14).setCellValue(1);
						moChaKa19482++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19483")) {
					row19483 = sheet19483.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19483.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19483.createCell((short) 1).setCellValue(1);
						row19483.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19483++;
						payMoney19483 = payMoney19483 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19483.createCell((short) 3).setCellValue(1);
						freeNum19483++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19483.createCell((short) 5).setCellValue(1);
						meiShi19483++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19483.createCell((short) 6).setCellValue(1);
						naTie19483++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19483.createCell((short) 7).setCellValue(1);
						moKa19483++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19483.createCell((short) 8).setCellValue(1);
						KaBu19483++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19483.createCell((short) 9).setCellValue(1);
						qiaoKe19483++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19483.createCell((short) 10).setCellValue(1);
						maQi19483++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19483.createCell((short) 11).setCellValue(1);
						niuNa19483++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19483.createCell((short) 12).setCellValue(1);
						qiaoKeNai19483++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19483.createCell((short) 13).setCellValue(1);
						moCha19483++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19483.createCell((short) 14).setCellValue(1);
						moChaKa19483++;
						moChaKaTotal++;
					}

				}

				if (entry.getKey().equals("19484")) {
					row19484 = sheet19484.createRow((int) i + 1);
					Infos info = infos.get(i);
					row19484.createCell((short) 0).setCellValue(info.getDate());
					if (!info.getPayType().equals("提货码")) {
						row19484.createCell((short) 1).setCellValue(1);
						row19484.createCell((short) 2).setCellValue(info.getpMoney());
						payNum19484++;
						payMoney19484 = payMoney19484 + info.getpMoney();
						payNumTotal++;
						payMoneyTotal += info.getpMoney();
					} else {
						row19484.createCell((short) 3).setCellValue(1);
						freeNum19484++;
						freeNumTotal++;
					}

					if (info.getpName().equals("美式咖啡加糖")) {
						row19484.createCell((short) 5).setCellValue(1);
						meiShi19484++;
						meiShiTotal++;
					}
					if (info.getpName().equals("拿铁加糖")) {
						row19484.createCell((short) 6).setCellValue(1);
						naTie19484++;
						naTieTotal++;
					}
					if (info.getpName().equals("摩卡加糖")) {
						row19484.createCell((short) 7).setCellValue(1);
						moKa19484++;
						moKaTotal++;
					}
					if (info.getpName().equals("卡布奇诺加糖")) {
						row19484.createCell((short) 8).setCellValue(1);
						KaBu19484++;
						KaBuTotal++;
					}
					if (info.getpName().equals("巧克力")) {
						row19484.createCell((short) 9).setCellValue(1);
						qiaoKe19484++;
						qiaoKeTotal++;
					}

					if (info.getpName().equals("玛琪雅朵加糖")) {
						row19484.createCell((short) 10).setCellValue(1);
						maQi19484++;
						maQiTotal++;
					}
					if (info.getpName().equals("热牛奶加糖")) {
						row19484.createCell((short) 11).setCellValue(1);
						niuNa19484++;
						niuNaTotal++;
					}
					if (info.getpName().equals("巧克力牛奶")) {
						row19484.createCell((short) 12).setCellValue(1);
						qiaoKeNai19484++;
						qiaoKeNaiTotal++;
					}
					if (info.getpName().equals("抹茶")) {
						row19484.createCell((short) 13).setCellValue(1);
						moCha19484++;
						moChaTotal++;
					}
					if (info.getpName().equals("抹茶咖啡")) {
						row19484.createCell((short) 14).setCellValue(1);
						moChaKa19484++;
						moChaKaTotal++;
					}

				}

			}

			if (entry.getKey().equals("18640")) {
				row18640 = sheet18640.createRow((int) lineNum + 1);
				row18640.createCell((short) 0).setCellValue("总计");
				row18640.createCell((short) 1).setCellValue(payNum18640);
				row18640.createCell((short) 2).setCellValue(payMoney18640);
				row18640.createCell((short) 3).setCellValue(freeNum18640);
				row18640.createCell((short) 5).setCellValue(meiShi18640);
				row18640.createCell((short) 6).setCellValue(naTie18640);
				row18640.createCell((short) 7).setCellValue(moKa18640);
				row18640.createCell((short) 8).setCellValue(KaBu18640);
				row18640.createCell((short) 9).setCellValue(qiaoKe18640);
				row18640.createCell((short) 10).setCellValue(maQi18640);
				row18640.createCell((short) 11).setCellValue(niuNa18640);
				row18640.createCell((short) 12).setCellValue(qiaoKeNai18640);
				row18640.createCell((short) 13).setCellValue(moCha18640);
				row18640.createCell((short) 14).setCellValue(moChaKa18640);

				rowDaySum = sheetSum.createRow((int) 1);
				rowDaySum.createCell((short) 0).setCellValue("18640光耀东方总部");
				rowDaySum.createCell((short) 1).setCellValue(payNum18640);
				rowDaySum.createCell((short) 2).setCellValue(payMoney18640);
				rowDaySum.createCell((short) 3).setCellValue(freeNum18640);
				rowDaySum.createCell((short) 5).setCellValue(meiShi18640);
				rowDaySum.createCell((short) 6).setCellValue(naTie18640);
				rowDaySum.createCell((short) 7).setCellValue(moKa18640);
				rowDaySum.createCell((short) 8).setCellValue(KaBu18640);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe18640);
				rowDaySum.createCell((short) 10).setCellValue(maQi18640);
				rowDaySum.createCell((short) 11).setCellValue(niuNa18640);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai18640);
				rowDaySum.createCell((short) 13).setCellValue(moCha18640);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa18640);
			}

			if (entry.getKey().equals("19452")) {
				row19452 = sheet19452.createRow((int) lineNum + 1);
				row19452.createCell((short) 0).setCellValue("总计");
				row19452.createCell((short) 1).setCellValue(payNum19452);
				row19452.createCell((short) 2).setCellValue(payMoney19452);
				row19452.createCell((short) 3).setCellValue(freeNum19452);
				row19452.createCell((short) 5).setCellValue(meiShi19452);
				row19452.createCell((short) 6).setCellValue(naTie19452);
				row19452.createCell((short) 7).setCellValue(moKa19452);
				row19452.createCell((short) 8).setCellValue(KaBu19452);
				row19452.createCell((short) 9).setCellValue(qiaoKe19452);
				row19452.createCell((short) 10).setCellValue(maQi19452);
				row19452.createCell((short) 11).setCellValue(niuNa19452);
				row19452.createCell((short) 12).setCellValue(qiaoKeNai19452);
				row19452.createCell((short) 13).setCellValue(moCha19452);
				row19452.createCell((short) 14).setCellValue(moChaKa19452);

				rowDaySum = sheetSum.createRow((int) 2);
				rowDaySum.createCell((short) 0).setCellValue("19452光耀东方广场");
				rowDaySum.createCell((short) 1).setCellValue(payNum19452);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19452);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19452);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19452);
				rowDaySum.createCell((short) 6).setCellValue(naTie19452);
				rowDaySum.createCell((short) 7).setCellValue(moKa19452);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19452);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19452);
				rowDaySum.createCell((short) 10).setCellValue(maQi19452);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19452);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19452);
				rowDaySum.createCell((short) 13).setCellValue(moCha19452);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19452);
			}

			if (entry.getKey().equals("19453")) {
				row19453 = sheet19453.createRow((int) lineNum + 1);
				row19453.createCell((short) 0).setCellValue("总计");
				row19453.createCell((short) 1).setCellValue(payNum19453);
				row19453.createCell((short) 2).setCellValue(payMoney19453);
				row19453.createCell((short) 3).setCellValue(freeNum19453);
				row19453.createCell((short) 5).setCellValue(meiShi19453);
				row19453.createCell((short) 6).setCellValue(naTie19453);
				row19453.createCell((short) 7).setCellValue(moKa19453);
				row19453.createCell((short) 8).setCellValue(KaBu19453);
				row19453.createCell((short) 9).setCellValue(qiaoKe19453);
				row19453.createCell((short) 10).setCellValue(maQi19453);
				row19453.createCell((short) 11).setCellValue(niuNa19453);
				row19453.createCell((short) 12).setCellValue(qiaoKeNai19453);
				row19453.createCell((short) 13).setCellValue(moCha19453);
				row19453.createCell((short) 14).setCellValue(moChaKa19453);

				rowDaySum = sheetSum.createRow((int) 3);
				rowDaySum.createCell((short) 0).setCellValue("19453联动优势1");
				rowDaySum.createCell((short) 1).setCellValue(payNum19453);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19453);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19453);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19453);
				rowDaySum.createCell((short) 6).setCellValue(naTie19453);
				rowDaySum.createCell((short) 7).setCellValue(moKa19453);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19453);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19453);
				rowDaySum.createCell((short) 10).setCellValue(maQi19453);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19453);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19453);
				rowDaySum.createCell((short) 13).setCellValue(moCha19453);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19453);
			}

			if (entry.getKey().equals("19454")) {
				row19454 = sheet19454.createRow((int) lineNum + 1);
				row19454.createCell((short) 0).setCellValue("总计");
				row19454.createCell((short) 1).setCellValue(payNum19454);
				row19454.createCell((short) 2).setCellValue(payMoney19454);
				row19454.createCell((short) 3).setCellValue(freeNum19454);
				row19454.createCell((short) 5).setCellValue(meiShi19454);
				row19454.createCell((short) 6).setCellValue(naTie19454);
				row19454.createCell((short) 7).setCellValue(moKa19454);
				row19454.createCell((short) 8).setCellValue(KaBu19454);
				row19454.createCell((short) 9).setCellValue(qiaoKe19454);
				row19454.createCell((short) 10).setCellValue(maQi19454);
				row19454.createCell((short) 11).setCellValue(niuNa19454);
				row19454.createCell((short) 12).setCellValue(qiaoKeNai19454);
				row19454.createCell((short) 13).setCellValue(moCha19454);
				row19454.createCell((short) 14).setCellValue(moChaKa19454);

				rowDaySum = sheetSum.createRow((int) 4);
				rowDaySum.createCell((short) 0).setCellValue("19454联动优势9");
				rowDaySum.createCell((short) 1).setCellValue(payNum19454);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19454);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19454);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19454);
				rowDaySum.createCell((short) 6).setCellValue(naTie19454);
				rowDaySum.createCell((short) 7).setCellValue(moKa19454);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19454);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19454);
				rowDaySum.createCell((short) 10).setCellValue(maQi19454);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19454);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19454);
				rowDaySum.createCell((short) 13).setCellValue(moCha19454);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19454);
			}

			if (entry.getKey().equals("19455")) {
				row19455 = sheet19455.createRow((int) lineNum + 1);
				row19455.createCell((short) 0).setCellValue("总计");
				row19455.createCell((short) 1).setCellValue(payNum19455);
				row19455.createCell((short) 2).setCellValue(payMoney19455);
				row19455.createCell((short) 3).setCellValue(freeNum19455);
				row19455.createCell((short) 5).setCellValue(meiShi19455);
				row19455.createCell((short) 6).setCellValue(naTie19455);
				row19455.createCell((short) 7).setCellValue(moKa19455);
				row19455.createCell((short) 8).setCellValue(KaBu19455);
				row19455.createCell((short) 9).setCellValue(qiaoKe19455);
				row19455.createCell((short) 10).setCellValue(maQi19455);
				row19455.createCell((short) 11).setCellValue(niuNa19455);
				row19455.createCell((short) 12).setCellValue(qiaoKeNai19455);
				row19455.createCell((short) 13).setCellValue(moCha19455);
				row19455.createCell((short) 14).setCellValue(moChaKa19455);

				rowDaySum = sheetSum.createRow((int) 5);
				rowDaySum.createCell((short) 0).setCellValue("19455烽火科技");
				rowDaySum.createCell((short) 1).setCellValue(payNum19455);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19455);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19455);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19455);
				rowDaySum.createCell((short) 6).setCellValue(naTie19455);
				rowDaySum.createCell((short) 7).setCellValue(moKa19455);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19455);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19455);
				rowDaySum.createCell((short) 10).setCellValue(maQi19455);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19455);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19455);
				rowDaySum.createCell((short) 13).setCellValue(moCha19455);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19455);
			}

			if (entry.getKey().equals("19456")) {
				row19456 = sheet19456.createRow((int) lineNum + 1);
				row19456.createCell((short) 0).setCellValue("总计");
				row19456.createCell((short) 1).setCellValue(payNum19456);
				row19456.createCell((short) 2).setCellValue(payMoney19456);
				row19456.createCell((short) 3).setCellValue(freeNum19456);
				row19456.createCell((short) 5).setCellValue(meiShi19456);
				row19456.createCell((short) 6).setCellValue(naTie19456);
				row19456.createCell((short) 7).setCellValue(moKa19456);
				row19456.createCell((short) 8).setCellValue(KaBu19456);
				row19456.createCell((short) 9).setCellValue(qiaoKe19456);
				row19456.createCell((short) 10).setCellValue(maQi19456);
				row19456.createCell((short) 11).setCellValue(niuNa19456);
				row19456.createCell((short) 12).setCellValue(qiaoKeNai19456);
				row19456.createCell((short) 13).setCellValue(moCha19456);
				row19456.createCell((short) 14).setCellValue(moChaKa19456);

				rowDaySum = sheetSum.createRow((int) 6);
				rowDaySum.createCell((short) 0).setCellValue("19456房天下3");
				rowDaySum.createCell((short) 1).setCellValue(payNum19456);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19456);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19456);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19456);
				rowDaySum.createCell((short) 6).setCellValue(naTie19456);
				rowDaySum.createCell((short) 7).setCellValue(moKa19456);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19456);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19456);
				rowDaySum.createCell((short) 10).setCellValue(maQi19456);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19456);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19456);
				rowDaySum.createCell((short) 13).setCellValue(moCha19456);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19456);
			}

			if (entry.getKey().equals("19457")) {
				row19457 = sheet19457.createRow((int) lineNum + 1);
				row19457.createCell((short) 0).setCellValue("总计");
				row19457.createCell((short) 1).setCellValue(payNum19457);
				row19457.createCell((short) 2).setCellValue(payMoney19457);
				row19457.createCell((short) 3).setCellValue(freeNum19457);
				row19457.createCell((short) 5).setCellValue(meiShi19457);
				row19457.createCell((short) 6).setCellValue(naTie19457);
				row19457.createCell((short) 7).setCellValue(moKa19457);
				row19457.createCell((short) 8).setCellValue(KaBu19457);
				row19457.createCell((short) 9).setCellValue(qiaoKe19457);
				row19457.createCell((short) 10).setCellValue(maQi19457);
				row19457.createCell((short) 11).setCellValue(niuNa19457);
				row19457.createCell((short) 12).setCellValue(qiaoKeNai19457);
				row19457.createCell((short) 13).setCellValue(moCha19457);
				row19457.createCell((short) 14).setCellValue(moChaKa19457);

				rowDaySum = sheetSum.createRow((int) 7);
				rowDaySum.createCell((short) 0).setCellValue("19457房天下1");
				rowDaySum.createCell((short) 1).setCellValue(payNum19457);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19457);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19457);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19457);
				rowDaySum.createCell((short) 6).setCellValue(naTie19457);
				rowDaySum.createCell((short) 7).setCellValue(moKa19457);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19457);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19457);
				rowDaySum.createCell((short) 10).setCellValue(maQi19457);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19457);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19457);
				rowDaySum.createCell((short) 13).setCellValue(moCha19457);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19457);
			}

			if (entry.getKey().equals("19458")) {
				row19458 = sheet19458.createRow((int) lineNum + 1);
				row19458.createCell((short) 0).setCellValue("总计");
				row19458.createCell((short) 1).setCellValue(payNum19458);
				row19458.createCell((short) 2).setCellValue(payMoney19458);
				row19458.createCell((short) 3).setCellValue(freeNum19458);
				row19458.createCell((short) 5).setCellValue(meiShi19458);
				row19458.createCell((short) 6).setCellValue(naTie19458);
				row19458.createCell((short) 7).setCellValue(moKa19458);
				row19458.createCell((short) 8).setCellValue(KaBu19458);
				row19458.createCell((short) 9).setCellValue(qiaoKe19458);
				row19458.createCell((short) 10).setCellValue(maQi19458);
				row19458.createCell((short) 11).setCellValue(niuNa19458);
				row19458.createCell((short) 12).setCellValue(qiaoKeNai19458);
				row19458.createCell((short) 13).setCellValue(moCha19458);
				row19458.createCell((short) 14).setCellValue(moChaKa19458);

				rowDaySum = sheetSum.createRow((int) 8);
				rowDaySum.createCell((short) 0).setCellValue("19458H3C");
				rowDaySum.createCell((short) 1).setCellValue(payNum19458);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19458);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19458);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19458);
				rowDaySum.createCell((short) 6).setCellValue(naTie19458);
				rowDaySum.createCell((short) 7).setCellValue(moKa19458);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19458);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19458);
				rowDaySum.createCell((short) 10).setCellValue(maQi19458);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19458);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19458);
				rowDaySum.createCell((short) 13).setCellValue(moCha19458);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19458);

			}

			if (entry.getKey().equals("19459")) {
				row19459 = sheet19459.createRow((int) lineNum + 1);
				row19459.createCell((short) 0).setCellValue("总计");
				row19459.createCell((short) 1).setCellValue(payNum19459);
				row19459.createCell((short) 2).setCellValue(payMoney19459);
				row19459.createCell((short) 3).setCellValue(freeNum19459);
				row19459.createCell((short) 5).setCellValue(meiShi19459);
				row19459.createCell((short) 6).setCellValue(naTie19459);
				row19459.createCell((short) 7).setCellValue(moKa19459);
				row19459.createCell((short) 8).setCellValue(KaBu19459);
				row19459.createCell((short) 9).setCellValue(qiaoKe19459);
				row19459.createCell((short) 10).setCellValue(maQi19459);
				row19459.createCell((short) 11).setCellValue(niuNa19459);
				row19459.createCell((short) 12).setCellValue(qiaoKeNai19459);
				row19459.createCell((short) 13).setCellValue(moCha19459);
				row19459.createCell((short) 14).setCellValue(moChaKa19459);

				rowDaySum = sheetSum.createRow((int) 9);
				rowDaySum.createCell((short) 0).setCellValue("19459爱农驿站22");
				rowDaySum.createCell((short) 1).setCellValue(payNum19459);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19459);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19459);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19459);
				rowDaySum.createCell((short) 6).setCellValue(naTie19459);
				rowDaySum.createCell((short) 7).setCellValue(moKa19459);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19459);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19459);
				rowDaySum.createCell((short) 10).setCellValue(maQi19459);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19459);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19459);
				rowDaySum.createCell((short) 13).setCellValue(moCha19459);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19459);
			}

			if (entry.getKey().equals("19472")) {
				row19472 = sheet19472.createRow((int) lineNum + 1);
				row19472.createCell((short) 0).setCellValue("总计");
				row19472.createCell((short) 1).setCellValue(payNum19472);
				row19472.createCell((short) 2).setCellValue(payMoney19472);
				row19472.createCell((short) 3).setCellValue(freeNum19472);
				row19472.createCell((short) 5).setCellValue(meiShi19472);
				row19472.createCell((short) 6).setCellValue(naTie19472);
				row19472.createCell((short) 7).setCellValue(moKa19472);
				row19472.createCell((short) 8).setCellValue(KaBu19472);
				row19472.createCell((short) 9).setCellValue(qiaoKe19472);
				row19472.createCell((short) 10).setCellValue(maQi19472);
				row19472.createCell((short) 11).setCellValue(niuNa19472);
				row19472.createCell((short) 12).setCellValue(qiaoKeNai19472);
				row19472.createCell((short) 13).setCellValue(moCha19472);
				row19472.createCell((short) 14).setCellValue(moChaKa19472);

				rowDaySum = sheetSum.createRow((int) 10);
				rowDaySum.createCell((short) 0).setCellValue("19472百融金服北楼20");
				rowDaySum.createCell((short) 1).setCellValue(payNum19472);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19472);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19472);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19472);
				rowDaySum.createCell((short) 6).setCellValue(naTie19472);
				rowDaySum.createCell((short) 7).setCellValue(moKa19472);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19472);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19472);
				rowDaySum.createCell((short) 10).setCellValue(maQi19472);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19472);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19472);
				rowDaySum.createCell((short) 13).setCellValue(moCha19472);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19472);
			}

			if (entry.getKey().equals("19473")) {
				row19473 = sheet19473.createRow((int) lineNum + 1);
				row19473.createCell((short) 0).setCellValue("总计");
				row19473.createCell((short) 1).setCellValue(payNum19473);
				row19473.createCell((short) 2).setCellValue(payMoney19473);
				row19473.createCell((short) 3).setCellValue(freeNum19473);
				row19473.createCell((short) 5).setCellValue(meiShi19473);
				row19473.createCell((short) 6).setCellValue(naTie19473);
				row19473.createCell((short) 7).setCellValue(moKa19473);
				row19473.createCell((short) 8).setCellValue(KaBu19473);
				row19473.createCell((short) 9).setCellValue(qiaoKe19473);
				row19473.createCell((short) 10).setCellValue(maQi19473);
				row19473.createCell((short) 11).setCellValue(niuNa19473);
				row19473.createCell((short) 12).setCellValue(qiaoKeNai19473);
				row19473.createCell((short) 13).setCellValue(moCha19473);
				row19473.createCell((short) 14).setCellValue(moChaKa19473);

				rowDaySum = sheetSum.createRow((int) 11);
				rowDaySum.createCell((short) 0).setCellValue("19473百融金服南楼2");
				rowDaySum.createCell((short) 1).setCellValue(payNum19473);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19473);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19473);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19473);
				rowDaySum.createCell((short) 6).setCellValue(naTie19473);
				rowDaySum.createCell((short) 7).setCellValue(moKa19473);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19473);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19473);
				rowDaySum.createCell((short) 10).setCellValue(maQi19473);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19473);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19473);
				rowDaySum.createCell((short) 13).setCellValue(moCha19473);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19473);
			}

			if (entry.getKey().equals("19475")) {
				row19475 = sheet19475.createRow((int) lineNum + 1);
				row19475.createCell((short) 0).setCellValue("总计");
				row19475.createCell((short) 1).setCellValue(payNum19475);
				row19475.createCell((short) 2).setCellValue(payMoney19475);
				row19475.createCell((short) 3).setCellValue(freeNum19475);
				row19475.createCell((short) 5).setCellValue(meiShi19475);
				row19475.createCell((short) 6).setCellValue(naTie19475);
				row19475.createCell((short) 7).setCellValue(moKa19475);
				row19475.createCell((short) 8).setCellValue(KaBu19475);
				row19475.createCell((short) 9).setCellValue(qiaoKe19475);
				row19475.createCell((short) 10).setCellValue(maQi19475);
				row19475.createCell((short) 11).setCellValue(niuNa19475);
				row19475.createCell((short) 12).setCellValue(qiaoKeNai19475);
				row19475.createCell((short) 13).setCellValue(moCha19475);
				row19475.createCell((short) 14).setCellValue(moChaKa19475);

				rowDaySum = sheetSum.createRow((int) 12);
				rowDaySum.createCell((short) 0).setCellValue("19475海象金服2");
				rowDaySum.createCell((short) 1).setCellValue(payNum19475);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19475);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19475);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19475);
				rowDaySum.createCell((short) 6).setCellValue(naTie19475);
				rowDaySum.createCell((short) 7).setCellValue(moKa19475);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19475);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19475);
				rowDaySum.createCell((short) 10).setCellValue(maQi19475);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19475);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19475);
				rowDaySum.createCell((short) 13).setCellValue(moCha19475);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19475);
			}

			if (entry.getKey().equals("19481")) {
				row19481 = sheet19481.createRow((int) lineNum + 1);
				row19481.createCell((short) 0).setCellValue("总计");
				row19481.createCell((short) 1).setCellValue(payNum19481);
				row19481.createCell((short) 2).setCellValue(payMoney19481);
				row19481.createCell((short) 3).setCellValue(freeNum19481);
				row19481.createCell((short) 5).setCellValue(meiShi19481);
				row19481.createCell((short) 6).setCellValue(naTie19481);
				row19481.createCell((short) 7).setCellValue(moKa19481);
				row19481.createCell((short) 8).setCellValue(KaBu19481);
				row19481.createCell((short) 9).setCellValue(qiaoKe19481);
				row19481.createCell((short) 10).setCellValue(maQi19481);
				row19481.createCell((short) 11).setCellValue(niuNa19481);
				row19481.createCell((short) 12).setCellValue(qiaoKeNai19481);
				row19481.createCell((short) 13).setCellValue(moCha19481);
				row19481.createCell((short) 14).setCellValue(moChaKa19481);

				rowDaySum = sheetSum.createRow((int) 13);
				rowDaySum.createCell((short) 0).setCellValue("19481云海肴");
				rowDaySum.createCell((short) 1).setCellValue(payNum19481);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19481);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19481);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19481);
				rowDaySum.createCell((short) 6).setCellValue(naTie19481);
				rowDaySum.createCell((short) 7).setCellValue(moKa19481);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19481);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19481);
				rowDaySum.createCell((short) 10).setCellValue(maQi19481);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19481);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19481);
				rowDaySum.createCell((short) 13).setCellValue(moCha19481);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19481);
			}

			if (entry.getKey().equals("19482")) {
				row19482 = sheet19482.createRow((int) lineNum + 1);
				row19482.createCell((short) 0).setCellValue("总计");
				row19482.createCell((short) 1).setCellValue(payNum19482);
				row19482.createCell((short) 2).setCellValue(payMoney19482);
				row19482.createCell((short) 3).setCellValue(freeNum19482);
				row19482.createCell((short) 5).setCellValue(meiShi19482);
				row19482.createCell((short) 6).setCellValue(naTie19482);
				row19482.createCell((short) 7).setCellValue(moKa19482);
				row19482.createCell((short) 8).setCellValue(KaBu19482);
				row19482.createCell((short) 9).setCellValue(qiaoKe19482);
				row19482.createCell((short) 10).setCellValue(maQi19482);
				row19482.createCell((short) 11).setCellValue(niuNa19482);
				row19482.createCell((short) 12).setCellValue(qiaoKeNai19482);
				row19482.createCell((short) 13).setCellValue(moCha19482);
				row19482.createCell((short) 14).setCellValue(moChaKa19482);

				rowDaySum = sheetSum.createRow((int) 14);
				rowDaySum.createCell((short) 0).setCellValue("19482意锐新创D5");
				rowDaySum.createCell((short) 1).setCellValue(payNum19482);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19482);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19482);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19482);
				rowDaySum.createCell((short) 6).setCellValue(naTie19482);
				rowDaySum.createCell((short) 7).setCellValue(moKa19482);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19482);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19482);
				rowDaySum.createCell((short) 10).setCellValue(maQi19482);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19482);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19482);
				rowDaySum.createCell((short) 13).setCellValue(moCha19482);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19482);
			}

			if (entry.getKey().equals("19483")) {
				row19483 = sheet19483.createRow((int) lineNum + 1);
				row19483.createCell((short) 0).setCellValue("总计");
				row19483.createCell((short) 1).setCellValue(payNum19483);
				row19483.createCell((short) 2).setCellValue(payMoney19483);
				row19483.createCell((short) 3).setCellValue(freeNum19483);
				row19483.createCell((short) 5).setCellValue(meiShi19483);
				row19483.createCell((short) 6).setCellValue(naTie19483);
				row19483.createCell((short) 7).setCellValue(moKa19483);
				row19483.createCell((short) 8).setCellValue(KaBu19483);
				row19483.createCell((short) 9).setCellValue(qiaoKe19483);
				row19483.createCell((short) 10).setCellValue(maQi19483);
				row19483.createCell((short) 11).setCellValue(niuNa19483);
				row19483.createCell((short) 12).setCellValue(qiaoKeNai19483);
				row19483.createCell((short) 13).setCellValue(moCha19483);
				row19483.createCell((short) 14).setCellValue(moChaKa19483);

				rowDaySum = sheetSum.createRow((int) 15);
				rowDaySum.createCell((short) 0).setCellValue("19483中科金财");
				rowDaySum.createCell((short) 1).setCellValue(payNum19483);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19483);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19483);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19483);
				rowDaySum.createCell((short) 6).setCellValue(naTie19483);
				rowDaySum.createCell((short) 7).setCellValue(moKa19483);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19483);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19483);
				rowDaySum.createCell((short) 10).setCellValue(maQi19483);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19483);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19483);
				rowDaySum.createCell((short) 13).setCellValue(moCha19483);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19483);
			}

			if (entry.getKey().equals("19484")) {
				row19484 = sheet19484.createRow((int) lineNum + 1);
				row19484.createCell((short) 0).setCellValue("总计");
				row19484.createCell((short) 1).setCellValue(payNum19484);
				row19484.createCell((short) 2).setCellValue(payMoney19484);
				row19484.createCell((short) 3).setCellValue(freeNum19484);
				row19484.createCell((short) 5).setCellValue(meiShi19484);
				row19484.createCell((short) 6).setCellValue(naTie19484);
				row19484.createCell((short) 7).setCellValue(moKa19484);
				row19484.createCell((short) 8).setCellValue(KaBu19484);
				row19484.createCell((short) 9).setCellValue(qiaoKe19484);
				row19484.createCell((short) 10).setCellValue(maQi19484);
				row19484.createCell((short) 11).setCellValue(niuNa19484);
				row19484.createCell((short) 12).setCellValue(qiaoKeNai19484);
				row19484.createCell((short) 13).setCellValue(moCha19484);
				row19484.createCell((short) 14).setCellValue(moChaKa19484);

				rowDaySum = sheetSum.createRow((int) 16);
				rowDaySum.createCell((short) 0).setCellValue("19484丰瑞祥");
				rowDaySum.createCell((short) 1).setCellValue(payNum19484);
				rowDaySum.createCell((short) 2).setCellValue(payMoney19484);
				rowDaySum.createCell((short) 3).setCellValue(freeNum19484);
				rowDaySum.createCell((short) 5).setCellValue(meiShi19484);
				rowDaySum.createCell((short) 6).setCellValue(naTie19484);
				rowDaySum.createCell((short) 7).setCellValue(moKa19484);
				rowDaySum.createCell((short) 8).setCellValue(KaBu19484);
				rowDaySum.createCell((short) 9).setCellValue(qiaoKe19484);
				rowDaySum.createCell((short) 10).setCellValue(maQi19484);
				rowDaySum.createCell((short) 11).setCellValue(niuNa19484);
				rowDaySum.createCell((short) 12).setCellValue(qiaoKeNai19484);
				rowDaySum.createCell((short) 13).setCellValue(moCha19484);
				rowDaySum.createCell((short) 14).setCellValue(moChaKa19484);
			}

			rowDaySum = sheetSum.createRow((int) 17);
			rowDaySum.createCell((short) 0).setCellValue("总计");
			rowDaySum.createCell((short) 1).setCellValue(payNumTotal);
			rowDaySum.createCell((short) 2).setCellValue(payMoneyTotal);
			rowDaySum.createCell((short) 3).setCellValue(freeNumTotal);
			rowDaySum.createCell((short) 5).setCellValue(meiShiTotal);
			rowDaySum.createCell((short) 6).setCellValue(naTieTotal);
			rowDaySum.createCell((short) 7).setCellValue(moKaTotal);
			rowDaySum.createCell((short) 8).setCellValue(KaBuTotal);
			rowDaySum.createCell((short) 9).setCellValue(qiaoKeTotal);
			rowDaySum.createCell((short) 10).setCellValue(maQiTotal);
			rowDaySum.createCell((short) 11).setCellValue(niuNaTotal);
			rowDaySum.createCell((short) 12).setCellValue(qiaoKeNaiTotal);
			rowDaySum.createCell((short) 13).setCellValue(moChaTotal);
			rowDaySum.createCell((short) 14).setCellValue(moChaKaTotal);

		}

		// 第六步，将文件存到指定位置
		try {
			FileOutputStream fout = new FileOutputStream("D:/统计结果.xls");
			wb.write(fout);
			fout.close();
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
