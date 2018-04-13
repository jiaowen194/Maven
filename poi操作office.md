# poi操作office

## 导入包
* 导入poi只能操作excle 如果需要操作word就必须导入poi-scratchpad 还有apache的一些依赖包如：collections
```xml
		<dependency>
		    <groupId>org.apache.poi</groupId>
		    <artifactId>poi</artifactId>
		    <version>3.17</version>
		</dependency>
		<dependency>
		    <groupId>org.apache.poi</groupId>
		    <artifactId>poi-ooxml</artifactId>
		    <version>3.17</version>
		</dependency>
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-scratchpad</artifactId>
			<version>3.17</version>
		</dependency>
```

```下面的代码需要写入close方法```
## 操作excel

```java
	public static String WORD_MOULD_PATH = "E://officetest/"; //模板位置
	public static String WORD_MOULD_CONTRACT = "contract.doc"; //合同模板
	public static String WORD_EXPORT_PATH = "E://officetest/"; //生成的word存储位置
	public static String EXCEL_EXPORT_PATH = "E://officetest/"; //生成的excel存储位置
  
  	public static String wordExport(String wordName,HashMap<String, Object> map,String wordMould) {
		//map为空则报错
		if(map==null||map.size()==0) {
			throw new RuntimeException("生成excel:map没有数据");
		}
		
		//格式化文件名
		//查询文件名是否为空并为空赋值
		if (wordName==null||wordName=="") {
			wordName="未起文件名";
		}		
		//去掉文件名后缀并加上doc
		int i = wordName.lastIndexOf(".");
		wordName = wordName.substring(0, i);
		String fileName=wordName+".doc";
		String filePath = WORD_MOULD_PATH+wordMould;
				
		InputStream is =null; 
		HWPFDocument doc =null;
		OutputStream os = null;
		
		//导入word模板并修改文件变量
		try {
			is = new FileInputStream(filePath);
			doc = new HWPFDocument(is);
			Range range = doc.getRange();
			Set<String> keySet = map.keySet();
			for (String key : keySet) {
				String repKey = "${"+key+"}";
				range.replaceText(repKey, map.get(key).toString());
			}
			
			//文件写入
		
			os = new FileOutputStream(WORD_EXPORT_PATH+fileName);
			doc.write(os);
			
			System.out.println("Word文件生成成功...");  
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
		      closeStream(os);    
		      closeStream(is);   
		}
		return WORD_EXPORT_PATH+fileName; 
		
	}
	
  

```



## 操作word
```java
	public static String excelExport(String excelName,List<Map> dataList) {

		if(dataList==null||dataList.get(0)==null) {
			throw new RuntimeException("生成excel:dataList没有数据");
		}
		
		//格式化文件名
		//查询文件名是否为空并为空赋值
		if (excelName==null||excelName=="") {
			excelName="未起文件名";
		}		
		//去掉文件名后缀并加上xls
		int i = excelName.lastIndexOf(".");
		excelName = excelName.substring(0, i);
		String fileName=excelName+".xls";
		
		//创建excel结构
		HSSFWorkbook excl = null;//创建工作簿
		HSSFSheet sheet = null;//创建sheet  
		HSSFRow row = null;//创建行
		HSSFCell cell = null;//创建单元格
		
		//创建表结构
		excl = new HSSFWorkbook();	
		sheet =excl.createSheet("sheet1"); 
		
		//声明标题字体
		HSSFFont font=excl.createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		
		//单元格样式（表头）
		HSSFCellStyle cs=excl.createCellStyle();
		cs.setAlignment(HorizontalAlignment.CENTER); //居中
		cs.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
		cs.setFont(font); //设置字体
				
		//按列完成数据输出
		row = sheet.createRow(0);
		Map topMap = dataList.get(0);
		Set<String> keys = topMap.keySet();
		int keyNum = 0; 
		Object cellValue = null; //用来临时保存循环列表的单元格值
		for (String str : keys) {
			cell = row.createCell(keyNum);
			cellValue = str;
			cell.setCellValue(cellValue == null ? "" : cellValue.toString());
			cell.setCellStyle(cs); //给标题加样式
			keyNum++;
		}
		
		//循环数据集填写数据
		for (int j = 0; j < dataList.size(); j++) {
			row = sheet.createRow(j+1);
			Map dataMap = dataList.get(j);
			Collection<Object> values = dataMap.values();
			int valueNum=0;
			for (Object obj : values) {
				cell = row.createCell(valueNum);
				cellValue = obj;
				cell.setCellValue(cellValue == null ? "" : cellValue.toString());
				valueNum++;
			}			
		}
				
		//文件写入
		File file = new File("E://officetest/"+fileName);
		FileOutputStream fileOut =null;
		try {
			fileOut = new FileOutputStream(file);
			excl.write(fileOut);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			closeStream(fileOut);
		}
		System.out.println("Excel文件生成成功...");
		return EXCEL_EXPORT_PATH+fileName;
		
	}



```
