package com.kmap.util;

import java.awt.Color;
import java.io.Closeable;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.view.AbstractView;

import com.ibm.icu.math.BigDecimal;

import egovframework.rte.psl.dataaccess.util.EgovMap;

/*
 * servlet.xml에 bean추가 하여 사용
 * <bean id="excelDownloadView" class="com.kmap.util.ExcelDownloadView"/>
 * */
public class ExcelDownloadView extends AbstractView {
	 @SuppressWarnings("unchecked")
	@Override
	    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response)
	            throws Exception {
	        System.out.println("엑셀다운로드뷰클래스 진입");
	        Locale locale = (Locale) model.get("locale");
	        String workbookName = (String) model.get("workbookName");
	        System.out.println("엑셀다운로드뷰클래스 워크북이름 : "+workbookName);
	        // 겹치는 파일 이름 중복을 피하기 위해 시간을 이용해서 파일 이름에 추가
	        Date date = new Date();
	        SimpleDateFormat dayformat = new SimpleDateFormat("yyyyMMdd", locale);
	        SimpleDateFormat hourformat = new SimpleDateFormat("hhmmss", locale);
	        String day = dayformat.format(date);
	        String hour = hourformat.format(date);
	        String fileName = workbookName + "_" + day + "_" + hour + ".xlsx";         
	        
	        // 각 브라우저에 따른 파일이름 인코딩작업
	        fileName = this.fileNameBrowserEncoding(fileName, request);
	        
	        //response.setContentType("application/download;charset=utf-8");
	        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
	        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
	        response.setHeader("Content-Transfer-Encoding", "binary");
	        
	       OutputStream os = null;
	       XSSFWorkbook workbook = null;
	       
	       try {
	    	   System.out.println("try 진입");
	           //workbook = (XSSFWorkbook) model.get("workbook");
	           workbook = this.makeWorkbook( (String[]) model.get("excelHead"), (List<EgovMap>) model.get("resultList"));
	           os = response.getOutputStream();
	           System.out.println(workbook.getSheetName(0));
	           // 파일생성
	           workbook.write(os);
	       }catch (Exception e) {
	           e.printStackTrace();
	       } finally {
	           /* 자바 버전 7이상 AutoCloseable 사용 객체는 자동으로 close()적용
	            * XSSFWorkbook 객체는 AutoCloseable 사용함으로 자동 close()가 적용
	            * 자바 버전 6이하로 낮출 경우 poi 3.8 버전을 올려 close() 사용이 되어야함.
	            * */
	    	   /*if(workbook != null) {
	               try {
	                   workbook.close();
	               } catch (Exception e) {
	                   e.printStackTrace();
	               }
	           }*/
	           
	           if(os != null) {
	               try {
	                   os.close();
	               } catch (Exception e) {
	                   e.printStackTrace();
	               }
	           }
	       }
	    }
	 
	 //20210503 추가
	 protected XSSFWorkbook makeWorkbook(String[] excelHead, List<EgovMap> resultList)
	 	throws Exception {
		//resultList를 엑셀데이터로 변환
		  XSSFWorkbook workbook = new XSSFWorkbook();
		  
		  //엑셀 스타일
		  XSSFCellStyle style = workbook.createCellStyle();
		  //배경색 설정
		  style.setFillForegroundColor(new XSSFColor(Color.lightGray));//밝은 회색
		  style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		  //테두리 설정
		  style.setBorderBottom(BorderStyle.THICK);//굵은선
		  style.setBorderLeft(BorderStyle.THIN);//얇은선
		  style.setBorderRight(BorderStyle.THIN);//얇은선
		  style.setBorderTop(BorderStyle.THIN);//얇은선
		  //글정렬 설정
		  style.setAlignment(HorizontalAlignment.CENTER);//가운데 정렬
		  
		  //시트생성
		  XSSFSheet sheet = workbook.createSheet("점용허가");
		  
		  // 데이터가 없을 때
	      if(resultList.size() == 0) {
	      	Row row;
	        Cell cell;
	          
	        // 행 생성
	      	row = sheet.createRow(0);
	      	//시트 열 너비 설정
	  		sheet.setColumnWidth(0, 3000);
	  		// 해당 행의 열 셀 생성
	  		cell = row.createCell(0);
	  		// 내용 설정
	  		cell.setCellValue("데이터가 없습니다");
	      } else { //데이터가 존재할 때
	    	  Row row;
		        Cell cell;
		        for(int i = 0; i < resultList.size()+1; i++) {
		        	// 행 생성
		        	row = sheet.createRow(i);
		        	
			        	for(int j = 0; j <excelHead.length; j++) {
			        		//시트 열 너비 설정
			        		sheet.setColumnWidth(j, 5000);
			        		// 해당 행의 열 셀 생성
			        		cell = row.createCell(j);
			        		if(i == 0){
			        			cell.setCellValue(excelHead[j]); //엑셀 첫번째 줄 이름 설정
			        			sheet.getRow(i).getCell(j).setCellStyle(style); //엑셀 첫번째 줄 스타일 설정
			        		} else {
			        			if(resultList.get(i-1).getValue(j) instanceof String) {
			        				cell.setCellValue((String)resultList.get(i-1).getValue(j));
			        			} else if(resultList.get(i-1).getValue(j) instanceof Double) {
			        				cell.setCellValue(Double.parseDouble(this.nullChangeZero(String.valueOf(resultList.get(i-1).getValue(j)))));
			        			} else if(resultList.get(i-1).getValue(j) instanceof BigDecimal) {
			        				cell.setCellValue(Double.parseDouble(this.nullChangeZero(String.valueOf(resultList.get(i-1).getValue(j)))));
			        			} else {
			        				cell.setCellValue(this.nullChange(String.valueOf(resultList.get(i-1).getValue(j))));
			        			}
			        		}
			        	}
		        }
	      }
		
		return workbook;
	 }
	 
	 // 파일 이름 한글 인코딩 메소드
	 protected String fileNameBrowserEncoding(String fileName,  HttpServletRequest request) throws Exception {
		 
		 String browser = request.getHeader("User-Agent");
	        if (browser.indexOf("MSIE") > -1) {
	            fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20");
	        } else if (browser.indexOf("Trident") > -1) {       // IE11
	            fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20");
	        } else if (browser.indexOf("Firefox") > -1) {
	            fileName = "\"" + new String(fileName.getBytes("UTF-8"), "8859_1") + "\"";
	        } else if (browser.indexOf("Opera") > -1) {
	            fileName = "\"" + new String(fileName.getBytes("UTF-8"), "8859_1") + "\"";
	        } else if (browser.indexOf("Chrome") > -1) {
	            StringBuffer sb = new StringBuffer();
	            for (int i = 0; i < fileName.length(); i++) {
	               char c = fileName.charAt(i);
	               if (c > '~') {
	                     sb.append(URLEncoder.encode("" + c, "UTF-8"));
	                       } else {
	                             sb.append(c);
	                       }
	                }
	                fileName = sb.toString();
	        } else if (browser.indexOf("Safari") > -1){
	            fileName = "\"" + new String(fileName.getBytes("UTF-8"), "8859_1")+ "\"";
	        } else {
	             fileName = "\"" + new String(fileName.getBytes("UTF-8"), "8859_1")+ "\"";
	        }
		 
		 return fileName;
	 }
	 
	////////////////////20210503 추가
	//null값 0으로 치환
	protected String nullChangeZero(String param) throws Exception {
	if (param.equals("null") || param.equals(null) || param.equals("")) {
	param="0";
	return param;
	} else {
	return param;
	}
	}
	
	//"null"이나 ""(공백)을 null로 치환
	protected String nullChange(String param) throws Exception {
	if (param.equals("null") || param.equals("")) {
	param=null;
	return param;
	} else {
	return param;
	}
	}
	///////////////////////////////
}
