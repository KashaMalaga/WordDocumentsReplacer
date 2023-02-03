import java.util.logging.Formatter;
import java.util.logging.LogRecord;

class TextAreaFormatter extends Formatter {
    @Override
    public String format(LogRecord record) {
        return formatMessage(record) + System.lineSeparator();
    }
}