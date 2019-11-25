Excel to JSON

The project takes a Excel file as input, it converts some of the columns into a json array and then stores it to a column of another file in the .csv format. It is a springboot project.

ExcelReader.java is a main file.

Following dependencies are added in pom.xml

1.org.apache.poi
2.com.googlecode.json-simple
3.com.opencsv (optional)
4.commons-io (optional)

Following is a sample output json:

[{"question":"Categories","answer":"POINT OFINTEREST"},{"question":"Sub Categories","answer":"Hospital/Polyclinic"},{"question":"Features","answer":"General"},{"question":"NAME","answer":"XYZ"},{"question":"LOCATION NAME","answer":"NA"},{"question":"TYPE","answer":"Hospital/Polyclinic"},{"question":"ADDITIONAL INFORMATION","answer":"NA,520007"},{"question":"SURVEY NO","answer":""}]
