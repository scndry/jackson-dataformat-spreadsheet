package io.github.scndry.jackson.dataformat.spreadsheet;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.function.BiFunction;

import org.apache.poi.ss.usermodel.DateUtil;

import com.fasterxml.jackson.databind.module.SimpleModule;

import io.github.scndry.jackson.dataformat.spreadsheet.deser.ExcelDateDeserializer;
import io.github.scndry.jackson.dataformat.spreadsheet.ser.ExcelDateSerializer;

/**
 * Jackson {@link SimpleModule} that registers serializers and
 * deserializers for converting between Java date/time types and
 * Excel date serial numbers using Apache POI's
 * {@link DateUtil}.
 *
 * @see SpreadsheetMapper
 */
public final class ExcelDateModule extends SimpleModule {

    public ExcelDateModule() {
        addExcelDateDeserializers();
        addExcelDateSerializers();
    }

    private void addExcelDateDeserializers() {
        final BiFunction<Double, Boolean, LocalDateTime> getLocalDateTime =
                DateUtil::getLocalDateTime;
        final BiFunction<Double, Boolean, LocalDate> getLocalDate =
                getLocalDateTime.andThen(LocalDateTime::toLocalDate);
        addDeserializer(Date.class, new ExcelDateDeserializer<>(DateUtil::getJavaDate));
        addDeserializer(Calendar.class, new ExcelDateDeserializer<>(DateUtil::getJavaCalendar));
        addDeserializer(LocalDate.class, new ExcelDateDeserializer<>(getLocalDate));
        addDeserializer(LocalDateTime.class, new ExcelDateDeserializer<>(getLocalDateTime));
    }

    private void addExcelDateSerializers() {
        addSerializer(Date.class, new ExcelDateSerializer<>(DateUtil::getExcelDate));
        addSerializer(Calendar.class, new ExcelDateSerializer<>(DateUtil::getExcelDate));
        addSerializer(LocalDate.class, new ExcelDateSerializer<>(DateUtil::getExcelDate));
        addSerializer(LocalDateTime.class, new ExcelDateSerializer<>(DateUtil::getExcelDate));
    }
}
