public class ExtentReportManager {
    private static ExtentReports extent;
    private static ExtentTest test;
    
    public static void initReport() {
        extent = new ExtentReports();
        ExtentSparkReporter spark = new ExtentSparkReporter("test-output/ExtentReport.html");
        
        // Configure to show scrollbars
        spark.config().setCss(
            "body { overflow-x: auto !important; overflow-y: auto !important; } " +
            ".test-content { overflow-x: auto !important; } " +
            ".log-container { overflow-x: auto !important; width: 100%; } " +
            "pre { white-space: pre-wrap; word-wrap: break-word; overflow-x: auto !important; } " +
            ".step-details { overflow-x: auto !important; }"
        );
        
        spark.config().setDocumentTitle("Test Automation Report");
        spark.config().setReportName("Regression Test Report");
        
        extent.attachReporter(spark);
    }
    
    public static void flushReport() {
        extent.flush();
    }
}