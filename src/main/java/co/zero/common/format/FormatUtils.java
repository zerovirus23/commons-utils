package co.zero.common.format;

import co.zero.common.Constants;

import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Currency;
import java.util.Date;

public class FormatUtils {
    private FormatUtils(){
    }

    /**
     * Get the date string representation according with the given pattern
     * @param date The date to format
     * @return The date formatted as string
     */
    public static String formatDate(Date date){
        return formatDate(date, Constants.DEFAULT_DATE_FORMAT);
    }

    /**
     * Get the date string representation according with the given pattern
     * @param date The date to format
     * @param pattern The format to apply
     * @return The date formatted as string
     */
    public static String formatDate(Date date, String pattern){
        DateFormat format = new SimpleDateFormat(pattern);
        return format.format(date);
    }

    /**
     * Get a number with max and min decimals
     * @param number The number to format
     * @param minDecimals Minimum decimals to show
     * @param maxDecimals Maximum decimals to show
     * @return The number formatted
     */
    public static String formatDouble(double number, int minDecimals, int maxDecimals){
        NumberFormat format = NumberFormat.getNumberInstance();
        format.setMinimumFractionDigits(minDecimals);
        format.setMaximumFractionDigits(maxDecimals);
        return format.format(number);
    }

    /**
     * Get a number with the default min and max decimals
     * @param number The  number to format
     * @return The number formatted with default values
     */
    public static String formatDouble(double number){
        return formatDouble(number, Constants.DEFAULT_MIN_DECIMALS, Constants.DEFAULT_MAX_DECIMALS);
    }

    /**
     * Get the representation of a number as a currency
     * @param value The number to format
     * @param currencyCode The ISO-4217 format
     * @return The number formatted as a currency
     */
    public static String formatCurrency(double value, String currencyCode){
        NumberFormat format = NumberFormat.getCurrencyInstance();
        Currency currency = Currency.getInstance(currencyCode);
        format.setMinimumFractionDigits(Constants.DEFAULT_MIN_DECIMALS);
        format.setMaximumFractionDigits(Constants.DEFAULT_MAX_DECIMALS);
        format.setCurrency(currency);
        return format.format(value);
    }

    /**
     * Get the Date object from a string given a format
     * @param dateFormatted The string containing the date
     * @param format Format of the date that has the string
     * @return The date object according to the format
     * @throws ParseException If the pattern of the date are wrong
     */
    public static Date parseDate(String dateFormatted, String format) throws ParseException {
        DateFormat formatter = new SimpleDateFormat(format);
        return formatter.parse(dateFormatted);
    }

    /**
     * Get the Date object from a string given a default format (yyyy-MM-dd)
     * @param dateFormatted The string containing the date
     * @return The date object according to the format
     * @throws ParseException If the pattern of the date are wrong
     */
    public static Date parseDate(String dateFormatted) throws ParseException {
        return parseDate(dateFormatted, Constants.DEFAULT_DATE_FORMAT);
    }
}