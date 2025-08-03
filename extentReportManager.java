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
        
        spark.config().setCss(
            // Fix category page container
            ".category-view { overflow-y: auto !important; max-height: 100vh !important; }" +
            ".category-container { overflow-y: auto !important; height: calc(100vh - 150px) !important; }" +
            ".categories-content { overflow-y: auto !important; }" +
            
            // Fix main content areas
            ".test-content { overflow-y: auto !important; max-height: 70vh !important; }" +
            ".test-list { overflow-y: auto !important; max-height: 70vh !important; }" +
            
            // Ensure body and html allow scrolling
            "html, body { overflow-y: auto !important; height: 100% !important; }" +
            
            // Fix specific category elements
            ".category-list { overflow-y: auto !important; max-height: 500px !important; }" +
            ".category-tests { overflow-y: auto !important; max-height: 400px !important; }" +
            
            // Force scrollbars to always show
            "::-webkit-scrollbar { width: 12px !important; }" +
            "::-webkit-scrollbar-track { background: #f1f1f1 !important; }" +
            "::-webkit-scrollbar-thumb { background: #888 !important; border-radius: 6px !important; }" +
            "::-webkit-scrollbar-thumb:hover { background: #555 !important; }"
        );
        

        spark.config().setDocumentTitle("Test Automation Report");
        spark.config().setReportName("Regression Test Report");
        
        extent.attachReporter(spark);
    }
    
    public static void flushReport() {
        extent.flush();
    }
}