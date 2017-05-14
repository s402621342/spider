package szw.Spider.Spider;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.rmi.CORBA.Util;

import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	String code;
    	int number=0;
    	for(int i=1;i<605000;i++){
    		if(i>3000&&i<300000){
    			continue;
    		}
    		if(i>300700&&i<600000){
    			continue;
    		}
    		code=String.valueOf(i);
    		while(code.length()<6){
    			code="0"+code;
    		}
    		String url="http://money.finance.sina.com.cn/corp/view/vFD_FinancialGuideLineHistory.php?stockid="+code+"&typecode=financialratios62";
        	String date="2016-12-31";
        	String[] temp=date.split("-");
        	String year=temp[0];
            String answer=createhttpClient(url, "");
            String name=null;
            String value=null;
            Pattern p=Pattern.compile("<h1 id=\"stockName\">(.+?)<span>");
            Matcher m = p.matcher(answer);
            if(m.find()) {
                name=m.group(1);
            }
            p = Pattern.compile("value='(.+?)' hoverText='"+date+"'/>");
            m = p.matcher(answer);
            if(m.find()){
            	value=m.group(1);
            }
            if(name!=null&&value!=null){
            	number++;
            	System.out.print(code+" ");
    	        System.out.print(name+" ");
    	        System.out.println(value);
           
            try{
            	InputStream instream = new FileInputStream("财务报表-净资产收益率2010.xls");   
            	HSSFWorkbook hssfWorkbook=new  HSSFWorkbook(instream);
            	for(int sheetnumber=0;sheetnumber<hssfWorkbook.getNumberOfSheets();sheetnumber++){
            		if(hssfWorkbook.getSheetName(sheetnumber).equals(year)){		//这张表
            			HSSFSheet hssfSheet=hssfWorkbook.getSheetAt(sheetnumber);
            			int yearnumber=0;			//对应年份所在的位置
            			HSSFRow hrow=hssfSheet.getRow(0);
            			for(int j=0;j<hrow.getLastCellNum();j++){				//找到年份所在那一列
            				try {
            					if(Double.parseDouble(year)==hrow.getCell(j).getNumericCellValue()){
                					yearnumber=j;
                					break;
                				}
							} catch (Exception e) {
								// TODO: handle exception
							}
            				
            			}
            			int startrow = 1;
            			boolean flag=false;
            			for(int row=startrow;row<hssfSheet.getLastRowNum();row++){
            				HSSFRow hssfRow=hssfSheet.getRow(row);
            				HSSFCell cell=hssfRow.getCell(0);
            				
            				try{
	            				if(i==(int)cell.getNumericCellValue()){
	            					HSSFCell cell2=hssfRow.getCell(yearnumber);
	            					if(cell2==null){
	            						cell2=hssfRow.createCell(yearnumber);
	            					}
	            					cell2.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
	            					cell2.setCellValue(Double.parseDouble(value));
	            					cell2=hssfRow.getCell(1);
	            					cell2.setCellValue(name);
	            					flag=true;
	            					startrow=row;
	            					break;
	            				}
            				}catch(Exception e){
            					e.printStackTrace();
            					
            				}
            			}
            			if(!flag){		//没有这行（新股票）
            				hssfSheet.shiftRows(number, hssfSheet.getLastRowNum(), 1);
                			HSSFRow row2=hssfSheet.createRow(number);
        					HSSFCell cell2=row2.createCell(0);
        					cell2.setCellValue(i);
        					cell2=row2.createCell(1);
        					cell2.setCellValue(name);
        					cell2=row2.createCell(yearnumber);
        					cell2.setCellValue(Double.parseDouble(value));
            			}
            			
            			
            		}
            	}
            	OutputStream os=new FileOutputStream("财务报表-净资产收益率2010.xls");  
            	hssfWorkbook.write(os);
            	os.close();
            	instream.close();
            	hssfWorkbook.close();
            }catch(Exception e){
            	e.printStackTrace();
            }
           }
    	}
    	
       
    }
    public static String createhttpClient(String url, String param) {
    	  HttpClient client = new HttpClient();
    	  String response = null;
    	  String keyword = null;
    	  PostMethod postMethod = new PostMethod(url);
    	//  try {
    	//   if (param != null)
//    	    keyword = new String(param.getBytes("gb2312"), "ISO-8859-1");
    	//  } catch (UnsupportedEncodingException e1) {
    	//   // TODO Auto-generated catch block
    	//   e1.printStackTrace();
    	//  }
    	  // NameValuePair[] data = { new NameValuePair("keyword", keyword) };
    	  // // 将表单的值放入postMethod中
    	  // postMethod.setRequestBody(data);
    	  // 以上部分是带参数抓取,我自己把它注销了．大家可以把注销消掉研究下
    	  try {
    	   int statusCode = client.executeMethod(postMethod);
    	   response = new String(postMethod.getResponseBodyAsString()
    	     .getBytes("ISO-8859-1"), "gb2312");
    	     //这里要注意下 gb2312要和你抓取网页的编码要一样
    	   String p = response.replaceAll("//&[a-zA-Z]{1,10};", "")
    	     .replaceAll("<[^>]*>", "");//去掉网页中带有html语言的标签
//    	   System.out.println(p);
    	  } catch (Exception e) {
    	   e.printStackTrace();
    	  }
    	  
    	  return response;
    	}

}
