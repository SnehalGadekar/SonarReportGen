package com.sonar.report;

import java.io.FileOutputStream;
import java.util.LinkedList;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.sonar.wsclient.SonarClient;
import org.sonar.wsclient.base.Paging;
import org.sonar.wsclient.issue.Issue;
import org.sonar.wsclient.issue.IssueClient;
import org.sonar.wsclient.issue.IssueQuery;
import org.sonar.wsclient.issue.Issues;

public class SonarXLSReport {

public static void main(String[] args) {
String login = "admin";
String password = "admin";

SonarClient client = SonarClient.create("http://35.154.168.177:9000");
client.builder().login(login);
client.builder().password(password);
IssueClient issueClient = client.issueClient();

IssueQuery query = IssueQuery.create();
query.severities("BLOCKER","CRITICAL","MAJOR","INFO","MINOR");
//query.severities("INFO");
query.resolved(false);

List<Issue> issues =new LinkedList<Issue>();
        Issues result;
        
        int pageIndex=1;
        do{
            query.pageIndex(pageIndex);
       //   query.pageSize(1000);
            result = issueClient.find(query);
            for(Issue issue:result.list()) 
			{
		if(issue.projectKey().equals(args[0]))
            	{
              issues.add(issue);
		}
            }
            pageIndex++;
        }
        //while(issues.size() < result.paging().total());
          while(pageIndex<=100);
        
    //    System.out.println("result.paging().total() :"+result.paging().total());
	System.out.println("Total Rows :"+issues.size());
	
	createExcel(issues,args[0]);
}

private static void createExcel(List<Issue> issueList,String reportName) {

try {
String filename = "/opt/jenkins/sonarqube-6.7.5/reports/"+reportName+".xls";

HSSFWorkbook workbook = new HSSFWorkbook();
HSSFSheet sheet = workbook.createSheet("FirstSheet");

HSSFRow rowhead = sheet.createRow((short) 0);
/*rowhead.createCell(0).setCellValue("Project Key");
rowhead.createCell(1).setCellValue("Component");
rowhead.createCell(2).setCellValue("Line");
rowhead.createCell(3).setCellValue("Rule Key");
rowhead.createCell(4).setCellValue("Severity");
rowhead.createCell(5).setCellValue("Resolutions");
rowhead.createCell(6).setCellValue("Message");*/

rowhead.createCell(0).setCellValue("ProjectKey/JiraID");
rowhead.createCell(1).setCellValue("TestCaseID");
rowhead.createCell(2).setCellValue("Action");
rowhead.createCell(4).setCellValue("TestResult");
rowhead.createCell(5).setCellValue("Description");
rowhead.createCell(6).setCellValue("Reporter");
rowhead.createCell(7).setCellValue("Assignee");
rowhead.createCell(8).setCellValue("Status");
rowhead.createCell(9).setCellValue("Priority");
rowhead.createCell(10).setCellValue("Comment");
rowhead.createCell(11).setCellValue("Line");
rowhead.createCell(12).setCellValue("Rule Key");
rowhead.createCell(13).setCellValue("Severity");
rowhead.createCell(14).setCellValue("Resolutions");

for (int i = 0; i < issueList.size(); i++) {
HSSFRow row = sheet.createRow((short) i+1);
row.createCell(0).setCellValue(issueList.get(i).projectKey());
row.createCell(2).setCellValue(issueList.get(i).componentKey());
row.createCell(4).setCellValue("fail");
row.createCell(5).setCellValue("Create JIRA Defect");
row.createCell(10).setCellValue(issueList.get(i).message());
row.createCell(11).setCellValue(String.valueOf(issueList.get(i).line()));
row.createCell(12).setCellValue(issueList.get(i).ruleKey());
row.createCell(13).setCellValue(issueList.get(i).severity());
row.createCell(14).setCellValue(issueList.get(i).resolution());
/*row.createCell(0).setCellValue(issueList.get(i).projectKey());
row.createCell(1).setCellValue(issueList.get(i).componentKey());
row.createCell(2).setCellValue(
String.valueOf(issueList.get(i).line()));
row.createCell(3).setCellValue(issueList.get(i).ruleKey());
row.createCell(4).setCellValue(issueList.get(i).severity());
row.createCell(5).setCellValue(issueList.get(i).resolution());
row.createCell(6).setCellValue(issueList.get(i).message());*/
}

FileOutputStream fileOut = new FileOutputStream(filename);
workbook.write(fileOut);
fileOut.close();
System.out.println("Your excel file has been generated!");

} catch (Exception ex) {
System.out.println(ex);

}
}

}
