package crmtest表结构导出;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;

import com.alibaba.druid.pool.DruidDataSource;

public class Main {
	public static void main(String[] args) throws Exception {
		int count = 0;
		String aa = "";
		JdbcTemplate jt = getJdbcTemplate();
		List<Map<String, Object>> list = jt.queryForList("SELECT T1.TABLE_NAME,T2.COMMENTS FROM USER_TABLES T1 LEFT JOIN USER_TAB_COMMENTS T2 ON T1.TABLE_NAME = T2.TABLE_NAME ORDER BY TABLE_NAME");
//		List<Map<String, Object>> list = jt.queryForList("SELECT T1.TABLE_NAME,T2.COMMENTS FROM USER_TABLES T1 LEFT JOIN USER_TAB_COMMENTS T2 ON T1.TABLE_NAME = T2.TABLE_NAME "
//				+ " where ROWNUM<3 ORDER BY TABLE_NAME");
//		
//		for(Map<String, Object> map:list){
//			System.out.println(map);
//		}
		
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		Sheet sh = wb.createSheet();
		
		//设置列宽
		sh.setColumnWidth(0, 20 * 256);
		sh.setColumnWidth(1, 10 * 256);
		sh.setColumnWidth(2, 9 * 256);
		sh.setColumnWidth(3, 12 * 256);
		sh.setColumnWidth(4, 80 * 256);
		
		
		//单元格头样式
		CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //单元格边框
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        
        
        //单元格行样式
		CellStyle cellStyle1 = wb.createCellStyle();
        //单元格边框
        cellStyle1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle1.setWrapText(true);
        
        //表头样式
		CellStyle cellStyle2 = wb.createCellStyle();
		//背景色
        cellStyle2.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        cellStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //字体
        Font font = wb.createFont();    
        font.setFontName("宋体");
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        cellStyle2.setFont(font);
        //单元格边框
        cellStyle2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        
		
        
        
		for(int  i= 0; i < list.size(); i++){
			if(list.get(i).get("COMMENTS")==null){
				count ++;
				if(list.get(i).get("TABLE_NAME").toString().length()==4){
					
					aa += list.get(i).get("TABLE_NAME") + "、";
				}
			}
			System.out.println("第"+i+"个表开始导出,表名:"+list.get(i).get("TABLE_NAME").toString());
            Row row = sh.createRow(sh.getLastRowNum()+1);
            //合并单元格  row row col col 坐标点
            CellRangeAddress cra=new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 4);
            sh.addMergedRegion(cra);
            RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN, cra, sh, wb);
            RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN, cra, sh, wb);
            RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN, cra, sh, wb);
            RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN, cra, sh, wb);
            
            Cell cell = row.createCell(0);
            cell.setCellValue(list.get(i).get("TABLE_NAME").toString());
            cell.setCellStyle(cellStyle);
            
            //表中文说明
            Row row1 = sh.createRow(sh.getLastRowNum()+1);
            //合并单元格  row row col col 坐标点
            CellRangeAddress cra1=new CellRangeAddress(row1.getRowNum(), row1.getRowNum(), 0, 4);
            sh.addMergedRegion(cra1);
            RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN, cra1, sh, wb);
            RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN, cra1, sh, wb);
            RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN, cra1, sh, wb);
            RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN, cra1, sh, wb);
            Cell cell1 = row1.createCell(0);
            cell1.setCellValue(list.get(i).get("COMMENTS")==null?"":list.get(i).get("COMMENTS").toString());
            cell1.setCellStyle(cellStyle);
            
            //写入说明行
            Row row3 = sh.createRow(sh.getLastRowNum()+1);
            Cell cell41 = row3.createCell(0);
            cell41.setCellStyle(cellStyle2);
            cell41.setCellValue("列名");
            Cell cell42 = row3.createCell(1);
            cell42.setCellStyle(cellStyle2);
            cell42.setCellValue("数据类型");
            Cell cell43 = row3.createCell(2);
            cell43.setCellStyle(cellStyle2);
            cell43.setCellValue("数据大小");
            Cell cell44 = row3.createCell(3);
            cell44.setCellStyle(cellStyle2);
            cell44.setCellValue("是否可以为空");
            Cell cell45 = row3.createCell(4);
            cell45.setCellStyle(cellStyle2);
            cell45.setCellValue("中文说明");
            
            //查询列名及说明
            List<Map<String, Object>> list1 = 
            		jt.queryForList(" select t1.COLUMN_NAME,t1.DATA_TYPE,t1.DATA_LENGTH,t1.NULLABLE, t2.COMMENTS from user_tab_columns t1 join user_col_comments t2 on t1.COLUMN_NAME = t2.COLUMN_NAME and t1.TABLE_NAME = t2.TABLE_NAME"+
									" where t1.Table_Name='"+list.get(i).get("TABLE_NAME")+"'"+
									" ORDER BY t1.COLUMN_ID");
            for(int j=0;j<list1.size();j++){
            	Row row2 = sh.createRow(sh.getLastRowNum()+1);
            	Cell cell21 = row2.createCell(0);
            	cell21.setCellStyle(cellStyle1);
            	cell21.setCellValue(list1.get(j).get("COLUMN_NAME")==null?"":list1.get(j).get("COLUMN_NAME").toString());
            	Cell cell22 = row2.createCell(1);
            	cell22.setCellStyle(cellStyle1);
            	cell22.setCellValue(list1.get(j).get("DATA_TYPE")==null?"":list1.get(j).get("DATA_TYPE").toString());
            	Cell cell23 = row2.createCell(2);
            	cell23.setCellStyle(cellStyle1);
            	cell23.setCellValue(list1.get(j).get("DATA_LENGTH")==null?"":list1.get(j).get("DATA_LENGTH").toString());
            	Cell cell24 = row2.createCell(3);
            	cell24.setCellStyle(cellStyle1);
            	cell24.setCellValue(list1.get(j).get("NULLABLE")==null?"":list1.get(j).get("NULLABLE").toString());
            	Cell cell25 = row2.createCell(4);
            	cell25.setCellStyle(cellStyle1);
            	cell25.setCellValue(list1.get(j).get("COMMENTS")==null?"":list1.get(j).get("COMMENTS").toString());
            }
            sh.createRow(sh.getLastRowNum()+1);
            sh.createRow(sh.getLastRowNum()+1);
            
        }
		
		String filePath = "d:/crm数据字典.xlsx";
		FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
        wb.close();
        System.out.println(count);
        System.out.println(aa);
	}

	/**
	 * 初始化数据源
	 * 为什么要用数据源呢？
	 * 因为JdbcTemplate要用
	 */
	private static DruidDataSource dataSource = null;
	static {
		dataSource = new DruidDataSource();
        dataSource.setUrl("jdbc:oracle:thin:@192.168.1.21:1521:crmtest");
        dataSource.setUsername("sshhbj");//用户名
        dataSource.setPassword("sshhbj");//密码
	}
	/**
	 * 获取JdbcTemplate
	 * 为什么要用JdbcTemplate呢？
	 * 因为不用写jdbc
	 */
	private static JdbcTemplate getJdbcTemplate() {
		JdbcTemplate jt = new JdbcTemplate(dataSource);
		return jt;
	}
}
