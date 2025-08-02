import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DateFormatConverter {
    
    public static void main(String[] args) {
        // Example 1: Convert mm-dd-yy to yyyy-mm-dd
        convertSimpleDate();
        
        // Example 2: Convert dd/MM/yyyy HH:mm:SS.s to yyyy-mm-dd HH:mm (with time reset to 00:00)
        convertDateTimeToDateOnly();
        
        // Example 3: Convert various formats
        convertVariousFormats();
    }
    
    // Convert mm-dd-yy to yyyy-mm-dd
    public static void convertSimpleDate() {
        System.out.println("=== Converting mm-dd-yy to yyyy-mm-dd ===");
        
        String inputDate = "12-25-23";
        DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("MM-dd-yy");
        DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        
        LocalDate date = LocalDate.parse(inputDate, inputFormatter);
        String formattedDate = date.format(outputFormatter);
        
        System.out.println("Input: " + inputDate);
        System.out.println("Output: " + formattedDate);
        System.out.println();
    }
    
    // Convert dd/MM/yyyy HH:mm:SS.s to yyyy-mm-dd HH:mm with time reset to 00:00
    public static void convertDateTimeToDateOnly() {
        System.out.println("=== Converting dd/MM/yyyy HH:mm:SS.s to yyyy-MM-dd 00:00 ===");
        
        String inputDateTime = "25/12/2023 14:30:45.123";
        DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss.SSS");
        DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
        
        LocalDateTime dateTime = LocalDateTime.parse(inputDateTime, inputFormatter);
        // Reset time to 00:00
        LocalDateTime resetDateTime = dateTime.toLocalDate().atStartOfDay();
        String formattedDateTime = resetDateTime.format(outputFormatter);
        
        System.out.println("Input: " + inputDateTime);
        System.out.println("Output: " + formattedDateTime);
        System.out.println();
    }
    
    // Convert various input formats
    public static void convertVariousFormats() {
        System.out.println("=== Converting Various Formats ===");
        
        // Example inputs
        String[] inputs = {
            "01-15-24",           // mm-dd-yy
            "03/04/2023 16:45:30.500",  // dd/MM/yyyy HH:mm:SS.s
            "2023-12-01 09:30:15.750",  // yyyy-MM-dd HH:mm:SS.s
            "15/08/2024 23:59:59.999"   // dd/MM/yyyy HH:mm:SS.s
        };
        
        DateTimeFormatter[] inputFormatters = {
            DateTimeFormatter.ofPattern("MM-dd-yy"),
            DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss.SSS"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SSS"),
            DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss.SSS")
        };
        
        DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
        
        for (int i = 0; i < inputs.length; i++) {
            try {
                String result;
                if (i == 0) {
                    // For date-only input, parse as LocalDate and convert to LocalDateTime with 00:00
                    LocalDate date = LocalDate.parse(inputs[i], inputFormatters[i]);
                    LocalDateTime dateTime = date.atStartOfDay();
                    result = dateTime.format(outputFormatter);
                } else {
                    // For datetime input, parse and reset time to 00:00
                    LocalDateTime dateTime = LocalDateTime.parse(inputs[i], inputFormatters[i]);
                    LocalDateTime resetDateTime = dateTime.toLocalDate().atStartOfDay();
                    result = resetDateTime.format(outputFormatter);
                }
                
                System.out.println("Input:  " + inputs[i]);
                System.out.println("Output: " + result);
                System.out.println();
            } catch (Exception e) {
                System.out.println("Error parsing: " + inputs[i] + " - " + e.getMessage());
            }
        }
    }
    
    // Utility method for flexible conversion
    public static String convertDateFormat(String inputDate, String inputPattern, String outputPattern, boolean resetTime) {
        try {
            if (inputPattern.contains("HH") || inputPattern.contains("mm") || inputPattern.contains("ss")) {
                // Input has time components
                DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern(inputPattern);
                DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern(outputPattern);
                
                LocalDateTime dateTime = LocalDateTime.parse(inputDate, inputFormatter);
                
                if (resetTime) {
                    // Reset time to 00:00:00
                    dateTime = dateTime.toLocalDate().atStartOfDay();
                }
                
                return dateTime.format(outputFormatter);
            } else {
                // Input is date-only
                DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern(inputPattern);
                LocalDate date = LocalDate.parse(inputDate, inputFormatter);
                
                if (outputPattern.contains("HH") || outputPattern.contains("mm") || outputPattern.contains("ss")) {
                    // Output needs time, so convert to LocalDateTime with 00:00
                    DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern(outputPattern);
                    LocalDateTime dateTime = date.atStartOfDay();
                    return dateTime.format(outputFormatter);
                } else {
                    // Both input and output are date-only
                    DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern(outputPattern);
                    return date.format(outputFormatter);
                }
            }
        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }
}