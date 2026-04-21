package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.annotation.*;
import com.fasterxml.jackson.databind.MapperFeature;
import com.fasterxml.jackson.databind.PropertyNamingStrategies;
import com.fasterxml.jackson.databind.exc.InvalidDefinitionException;
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.fasterxml.jackson.databind.annotation.JsonNaming;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.databind.ser.FilterProvider;
import com.fasterxml.jackson.databind.ser.impl.SimpleBeanPropertyFilter;
import com.fasterxml.jackson.databind.ser.impl.SimpleFilterProvider;
import com.fasterxml.jackson.databind.util.StdConverter;

import java.util.HashMap;
import java.util.Map;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Verifies compatibility with common Jackson annotations.
 * Ordered by usage frequency. Documents which annotations
 * work and which structurally cannot.
 */
class JacksonAnnotationTest {

    SpreadsheetMapper mapper;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    // ----------------------------------------------------------------
    // @JsonProperty
    // ----------------------------------------------------------------

    @DataGrid
    static class WithJsonProperty {
        @JsonProperty("Product Name")
        public String name;
        @JsonProperty("Unit Price")
        public double price;

        public WithJsonProperty() {}
        public WithJsonProperty(String name, double price) {
            this.name = name;
            this.price = price;
        }
    }

    @Test
    void jsonProperty_renamesColumn() throws Exception {
        File file = tempFile("json-property.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithJsonProperty("Apple", 1.50)), WithJsonProperty.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("Product Name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("Unit Price");
        }

        List<WithJsonProperty> read = mapper.readValues(file, WithJsonProperty.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).price).isEqualTo(1.50);
    }

    @Test
    void schemaReflectsJacksonAnnotations() throws Exception {
        SpreadsheetSchema schema = mapper.sheetSchemaFor(WithJsonProperty.class);
        Iterator<io.github.scndry.jackson.dataformat.spreadsheet.schema.Column> it =
                schema.iterator();
        assertThat(it.next().getName()).isEqualTo("Product Name");
        assertThat(it.next().getName()).isEqualTo("Unit Price");
    }

    // ----------------------------------------------------------------
    // @JsonIgnore
    // ----------------------------------------------------------------

    @DataGrid
    static class WithJsonIgnore {
        public String name;
        @JsonIgnore
        public String internalCode;
        public int quantity;

        public WithJsonIgnore() {}
        public WithJsonIgnore(String name, String internalCode, int quantity) {
            this.name = name;
            this.internalCode = internalCode;
            this.quantity = quantity;
        }
    }

    @Test
    void jsonIgnore_excludesColumn() throws Exception {
        File file = tempFile("json-ignore.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithJsonIgnore("Apple", "INTERNAL-001", 10)),
                WithJsonIgnore.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("quantity");
            assertThat(header.getCell(2)).isNull();
        }

        List<WithJsonIgnore> read = mapper.readValues(file, WithJsonIgnore.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).internalCode).isNull();
        assertThat(read.get(0).quantity).isEqualTo(10);
    }

    // ----------------------------------------------------------------
    // @JsonIgnoreProperties
    // ----------------------------------------------------------------

    @DataGrid
    @JsonIgnoreProperties({"password", "secret"})
    static class WithIgnoreProperties {
        public String name;
        public String password;
        public String secret;
        public int age;

        public WithIgnoreProperties() {}
        public WithIgnoreProperties(String name, String password, String secret, int age) {
            this.name = name;
            this.password = password;
            this.secret = secret;
            this.age = age;
        }
    }

    @Test
    void jsonIgnoreProperties_excludesMultiple() throws Exception {
        File file = tempFile("ignore-props.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithIgnoreProperties("Alice", "pass123", "s3cret", 30)),
                WithIgnoreProperties.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("age");
            assertThat(header.getCell(2)).isNull();
        }

        List<WithIgnoreProperties> read = mapper.readValues(file, WithIgnoreProperties.class);
        assertThat(read.get(0).name).isEqualTo("Alice");
        assertThat(read.get(0).age).isEqualTo(30);
        assertThat(read.get(0).password).isNull();
        assertThat(read.get(0).secret).isNull();
    }

    // ----------------------------------------------------------------
    // @JsonCreator
    // ----------------------------------------------------------------

    @DataGrid
    static class Immutable {
        private final String name;
        private final int value;

        @JsonCreator
        public Immutable(
                @JsonProperty("name") String name,
                @JsonProperty("value") int value) {
            this.name = name;
            this.value = value;
        }

        public String getName() { return name; }
        public int getValue() { return value; }
    }

    @Test
    void jsonCreator_constructorDeserialization() throws Exception {
        File file = tempFile("json-creator.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new Immutable("test", 42)), Immutable.class);

        List<Immutable> read = mapper.readValues(file, Immutable.class);
        assertThat(read).hasSize(1);
        assertThat(read.get(0).getName()).isEqualTo("test");
        assertThat(read.get(0).getValue()).isEqualTo(42);
    }

    // ----------------------------------------------------------------
    // @JsonPropertyOrder
    // ----------------------------------------------------------------

    @DataGrid
    @JsonPropertyOrder({"price", "name", "quantity"})
    static class WithPropertyOrder {
        public String name;
        public int quantity;
        public double price;

        public WithPropertyOrder() {}
        public WithPropertyOrder(String name, int quantity, double price) {
            this.name = name;
            this.quantity = quantity;
            this.price = price;
        }
    }

    @Test
    void jsonPropertyOrder_controlsColumnOrder() throws Exception {
        File file = tempFile("property-order.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithPropertyOrder("Apple", 10, 1.50)),
                WithPropertyOrder.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("price");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("quantity");
        }

        List<WithPropertyOrder> read = mapper.readValues(file, WithPropertyOrder.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).price).isEqualTo(1.50);
        assertThat(read.get(0).quantity).isEqualTo(10);
    }

    // ----------------------------------------------------------------
    // @JsonInclude
    // ----------------------------------------------------------------

    @DataGrid
    @JsonInclude(JsonInclude.Include.NON_NULL)
    static class WithNonNull {
        public String required;
        public String optional;

        public WithNonNull() {}
        public WithNonNull(String required, String optional) {
            this.required = required;
            this.optional = optional;
        }
    }

    @Test
    void jsonInclude_nonNull_skipsNullCell() throws Exception {
        File file = tempFile("non-null.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithNonNull("present", null)),
                WithNonNull.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row dataRow = wb.getSheetAt(0).getRow(1);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("present");
            assertThat(dataRow.getCell(1)).isNull();
        }

        List<WithNonNull> read = mapper.readValues(file, WithNonNull.class);
        assertThat(read.get(0).required).isEqualTo("present");
        assertThat(read.get(0).optional).isNull();
    }

    // ----------------------------------------------------------------
    // @JsonNaming
    // ----------------------------------------------------------------

    @DataGrid
    @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
    static class WithNamingStrategy {
        public String firstName;
        public String lastName;
        public int totalCount;

        public WithNamingStrategy() {}
        public WithNamingStrategy(String firstName, String lastName, int totalCount) {
            this.firstName = firstName;
            this.lastName = lastName;
            this.totalCount = totalCount;
        }
    }

    @Test
    void jsonNaming_appliesStrategy() throws Exception {
        File file = tempFile("json-naming.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithNamingStrategy("John", "Doe", 5)),
                WithNamingStrategy.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("first_name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("last_name");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("total_count");
        }

        List<WithNamingStrategy> read = mapper.readValues(file, WithNamingStrategy.class);
        assertThat(read.get(0).firstName).isEqualTo("John");
        assertThat(read.get(0).totalCount).isEqualTo(5);
    }

    // ----------------------------------------------------------------
    // @JsonGetter / @JsonSetter
    // ----------------------------------------------------------------

    @DataGrid
    static class WithGetterSetter {
        private String value;
        private int count;

        public WithGetterSetter() {}
        public WithGetterSetter(String value, int count) {
            this.value = value;
            this.count = count;
        }

        @JsonGetter("display_name")
        public String getValue() { return value; }
        @JsonSetter("display_name")
        public void setValue(String value) { this.value = value; }

        public int getCount() { return count; }
        public void setCount(int count) { this.count = count; }
    }

    @Test
    void jsonGetterSetter_renamesColumn() throws Exception {
        File file = tempFile("getter-setter.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithGetterSetter("test", 5)),
                WithGetterSetter.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("count");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("display_name");
        }

        List<WithGetterSetter> read = mapper.readValues(file, WithGetterSetter.class);
        assertThat(read.get(0).getValue()).isEqualTo("test");
        assertThat(read.get(0).getCount()).isEqualTo(5);
    }

    // ----------------------------------------------------------------
    // @JsonAutoDetect
    // ----------------------------------------------------------------

    @DataGrid
    @JsonAutoDetect(fieldVisibility = JsonAutoDetect.Visibility.ANY)
    static class WithPrivateFields {
        private String secret;
        private int count;

        public WithPrivateFields() {}
        public WithPrivateFields(String secret, int count) {
            this.secret = secret;
            this.count = count;
        }
    }

    @DataGrid
    static class WithPrivateFieldsNoDetect {
        private String secret;
        private int count;

        public WithPrivateFieldsNoDetect() {}
        public WithPrivateFieldsNoDetect(String secret, int count) {
            this.secret = secret;
            this.count = count;
        }
    }

    @Test
    void jsonAutoDetect_privateFields() throws Exception {
        File file = tempFile("private-fields.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithPrivateFields("hidden", 42)),
                WithPrivateFields.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("secret");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("count");
        }

        List<WithPrivateFields> read = mapper.readValues(file, WithPrivateFields.class);
        assertThat(read.get(0).secret).isEqualTo("hidden");
        assertThat(read.get(0).count).isEqualTo(42);

        // Without @JsonAutoDetect: no visible properties → fail fast
        assertThatThrownBy(() -> mapper.sheetSchemaFor(WithPrivateFieldsNoDetect.class))
                .isInstanceOf(InvalidDefinitionException.class)
                .hasMessageContaining("no visible properties");
    }

    // ----------------------------------------------------------------
    // @JsonValue / @JsonCreator on enum
    // ----------------------------------------------------------------

    enum Status {
        ACTIVE("A"), INACTIVE("I"), PENDING("P");
        private final String code;
        Status(String code) { this.code = code; }
        @JsonValue
        public String getCode() { return code; }

        @JsonCreator
        public static Status fromCode(String code) {
            for (Status s : values()) {
                if (s.code.equals(code)) return s;
            }
            return null;
        }
    }

    @DataGrid
    static class WithEnumValue {
        public String name;
        public Status status;

        public WithEnumValue() {}
        public WithEnumValue(String name, Status status) {
            this.name = name;
            this.status = status;
        }
    }

    @Test
    void jsonValue_enumSerializesToCustomValue() throws Exception {
        File file = tempFile("enum-value.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithEnumValue("Alice", Status.ACTIVE),
                new WithEnumValue("Bob", Status.PENDING)),
                WithEnumValue.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("A");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("P");
        }

        List<WithEnumValue> read = mapper.readValues(file, WithEnumValue.class);
        assertThat(read.get(0).status).isEqualTo(Status.ACTIVE);
        assertThat(read.get(1).status).isEqualTo(Status.PENDING);
    }

    // ----------------------------------------------------------------
    // @JsonEnumDefaultValue
    // ----------------------------------------------------------------

    enum Priority {
        HIGH, MEDIUM, LOW,
        @JsonEnumDefaultValue UNKNOWN
    }

    @DataGrid
    static class WithEnumDefault {
        public String name;
        public Priority priority;

        public WithEnumDefault() {}
        public WithEnumDefault(String name, Priority priority) {
            this.name = name;
            this.priority = priority;
        }
    }

    @Test
    void jsonEnumDefaultValue_fallbackOnUnknown() throws Exception {
        File file = tempFile("enum-default.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("priority");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("task");
            data.createCell(1).setCellValue("CRITICAL"); // not in enum
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper enumMapper = SpreadsheetMapper.builder()
                .enable(com.fasterxml.jackson.databind.DeserializationFeature
                        .READ_UNKNOWN_ENUM_VALUES_USING_DEFAULT_VALUE)
                .build();
        List<WithEnumDefault> read = enumMapper.readValues(file, WithEnumDefault.class);
        assertThat(read.get(0).priority).isEqualTo(Priority.UNKNOWN);
    }

    // ----------------------------------------------------------------
    // @JsonSerialize / @JsonDeserialize
    // ----------------------------------------------------------------

    static class UpperCaseConverter extends StdConverter<String, String> {
        @Override
        public String convert(String value) {
            return value != null ? value.toUpperCase() : null;
        }
    }

    static class LowerCaseConverter extends StdConverter<String, String> {
        @Override
        public String convert(String value) {
            return value != null ? value.toLowerCase() : null;
        }
    }

    @DataGrid
    static class WithCustomConverter {
        @JsonSerialize(converter = UpperCaseConverter.class)
        @JsonDeserialize(converter = LowerCaseConverter.class)
        public String name;
        public int value;

        public WithCustomConverter() {}
        public WithCustomConverter(String name, int value) {
            this.name = name;
            this.value = value;
        }
    }

    @Test
    void jsonSerialize_customConverter() throws Exception {
        File file = tempFile("converter.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithCustomConverter("hello", 1)),
                WithCustomConverter.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0)
                    .getStringCellValue()).isEqualTo("HELLO");
        }

        List<WithCustomConverter> read = mapper.readValues(file, WithCustomConverter.class);
        assertThat(read.get(0).name).isEqualTo("hello");
    }

    // ----------------------------------------------------------------
    // @JsonFormat
    // ----------------------------------------------------------------

    @DataGrid
    static class WithJsonFormat {
        public String name;
        @JsonFormat(shape = JsonFormat.Shape.STRING)
        public int numericAsString;

        public WithJsonFormat() {}
        public WithJsonFormat(String name, int numericAsString) {
            this.name = name;
            this.numericAsString = numericAsString;
        }
    }

    @DataGrid
    static class WithoutJsonFormat {
        public String name;
        public int numericValue;

        public WithoutJsonFormat() {}
        public WithoutJsonFormat(String name, int numericValue) {
            this.name = name;
            this.numericValue = numericValue;
        }
    }

    @Test
    void jsonFormat_shapeString_vs_default() throws Exception {
        File fileFormatted = tempFile("json-format.xlsx");
        mapper.writeValue(fileFormatted, Arrays.asList(
                new WithJsonFormat("test", 42)), WithJsonFormat.class);

        File fileDefault = tempFile("no-format.xlsx");
        mapper.writeValue(fileDefault, Arrays.asList(
                new WithoutJsonFormat("test", 42)), WithoutJsonFormat.class);

        try (XSSFWorkbook wbFormatted = new XSSFWorkbook(fileFormatted);
             XSSFWorkbook wbDefault = new XSSFWorkbook(fileDefault)) {
            assertThat(wbFormatted.getSheetAt(0).getRow(1).getCell(1)
                    .getCellType()).isEqualTo(CellType.STRING);
            assertThat(wbFormatted.getSheetAt(0).getRow(1).getCell(1)
                    .getStringCellValue()).isEqualTo("42");

            assertThat(wbDefault.getSheetAt(0).getRow(1).getCell(1)
                    .getCellType()).isEqualTo(CellType.NUMERIC);
        }
    }

    // ----------------------------------------------------------------
    // @JsonUnwrapped
    // ----------------------------------------------------------------

    @DataGrid
    static class WithUnwrapped {
        public String name;
        @JsonUnwrapped
        public InnerValue inner;

        public WithUnwrapped() {}
        public WithUnwrapped(String name, InnerValue inner) {
            this.name = name;
            this.inner = inner;
        }
    }

    static class InnerValue {
        public int x;
        public int y;

        public InnerValue() {}
        public InnerValue(int x, int y) {
            this.x = x;
            this.y = y;
        }
    }

    @DataGrid
    static class WithoutUnwrapped {
        public String name;
        public InnerValue inner;

        public WithoutUnwrapped() {}
        public WithoutUnwrapped(String name, InnerValue inner) {
            this.name = name;
            this.inner = inner;
        }
    }

    @Test
    void libraryFlattening_nestedHeaderUsesPath() throws Exception {
        File file = tempFile("no-unwrap.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithoutUnwrapped("A", new InnerValue(1, 2))),
                WithoutUnwrapped.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("inner/x");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("inner/y");
        }

        List<WithoutUnwrapped> read = mapper.readValues(file, WithoutUnwrapped.class);
        assertThat(read.get(0).inner.x).isEqualTo(1);
        assertThat(read.get(0).inner.y).isEqualTo(2);
    }

    @Test
    void jsonUnwrapped_roundTrip() throws Exception {
        File file = tempFile("unwrapped.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithUnwrapped("A", new InnerValue(1, 2))),
                WithUnwrapped.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("x");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("y");
        }

        List<WithUnwrapped> read = mapper.readValues(file, WithUnwrapped.class);
        assertThat(read.get(0).name).isEqualTo("A");
        assertThat(read.get(0).inner.x).isEqualTo(1);
        assertThat(read.get(0).inner.y).isEqualTo(2);
    }

    // ----------------------------------------------------------------
    // @JsonIncludeProperties
    // ----------------------------------------------------------------

    @DataGrid
    @JsonIncludeProperties({"name", "price"})
    static class WithIncludeProperties {
        public String name;
        public int quantity;
        public double price;
        public String description;

        public WithIncludeProperties() {}
        public WithIncludeProperties(String name, int quantity, double price, String description) {
            this.name = name;
            this.quantity = quantity;
            this.price = price;
            this.description = description;
        }
    }

    @Test
    void jsonIncludeProperties_whitelistColumns() throws Exception {
        File file = tempFile("include-props.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithIncludeProperties("Apple", 10, 1.50, "A fruit")),
                WithIncludeProperties.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("price");
            assertThat(header.getCell(2)).isNull();
        }

        List<WithIncludeProperties> read = mapper.readValues(file, WithIncludeProperties.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).price).isEqualTo(1.50);
        assertThat(read.get(0).quantity).isEqualTo(0);
        assertThat(read.get(0).description).isNull();
    }

    // ----------------------------------------------------------------
    // @JsonFilter
    // ----------------------------------------------------------------

    @DataGrid
    @JsonFilter("columnFilter")
    static class WithFilter {
        public String name;
        public int quantity;
        public double price;
        public String notes;

        public WithFilter() {}
        public WithFilter(String name, int quantity, double price, String notes) {
            this.name = name;
            this.quantity = quantity;
            this.price = price;
            this.notes = notes;
        }
    }

    @Test
    void jsonFilter_programmaticExclusion() throws Exception {
        FilterProvider filters = new SimpleFilterProvider()
                .addFilter("columnFilter",
                        SimpleBeanPropertyFilter.serializeAllExcept("notes", "quantity"));

        SpreadsheetMapper filterMapper = SpreadsheetMapper.builder().build();
        filterMapper.setFilterProvider(filters);

        File file = tempFile("filter.xlsx");
        filterMapper.writeValue(file, Arrays.asList(
                new WithFilter("Apple", 10, 1.50, "fresh")),
                WithFilter.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("price");
            assertThat(header.getCell(2)).isNull();
        }

        List<WithFilter> read = filterMapper.readValues(file, WithFilter.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).price).isEqualTo(1.50);
    }

    // ----------------------------------------------------------------
    // @JsonIgnoreType
    // ----------------------------------------------------------------

    @JsonIgnoreType
    static class AuditInfo {
        public String createdBy;
        public String updatedBy;
    }

    @DataGrid
    static class WithIgnoreType {
        public String name;
        public AuditInfo audit;
        public int value;

        public WithIgnoreType() {}
        public WithIgnoreType(String name, AuditInfo audit, int value) {
            this.name = name;
            this.audit = audit;
            this.value = value;
        }
    }

    @Test
    void jsonIgnoreType_excludesNestedType() throws Exception {
        AuditInfo audit = new AuditInfo();
        audit.createdBy = "admin";
        audit.updatedBy = "admin";

        File file = tempFile("ignore-type.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithIgnoreType("item", audit, 10)),
                WithIgnoreType.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("value");
            assertThat(header.getCell(2)).isNull();
        }

        List<WithIgnoreType> read = mapper.readValues(file, WithIgnoreType.class);
        assertThat(read.get(0).name).isEqualTo("item");
        assertThat(read.get(0).value).isEqualTo(10);
        assertThat(read.get(0).audit).isNull();
    }

    // ----------------------------------------------------------------
    // @JsonView
    // ----------------------------------------------------------------

    static class Views {
        static class Summary {}
        static class Detail extends Summary {}
    }

    @DataGrid
    static class WithViews {
        @JsonView(Views.Summary.class)
        public String name;
        @JsonView(Views.Detail.class)
        public String description;
        @JsonView(Views.Summary.class)
        public int quantity;

        public WithViews() {}
        public WithViews(String name, String description, int quantity) {
            this.name = name;
            this.description = description;
            this.quantity = quantity;
        }
    }

    @Test
    void jsonView_viaSheetWriterFor() throws Exception {
        SpreadsheetMapper viewMapper = SpreadsheetMapper.builder()
                .configure(MapperFeature.DEFAULT_VIEW_INCLUSION, false)
                .build();

        File file = tempFile("view.xlsx");
        viewMapper.sheetWriterForWithView(WithViews.class, Views.Summary.class)
                .writeValue(file, Arrays.asList(
                        new WithViews("Apple", "A fruit", 10)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("quantity");
            assertThat(header.getCell(2)).isNull();
        }
    }

    @Test
    void jsonView_writerWithView_notSupported() {
        File file = tempFile("view-fail.xlsx");
        assertThatThrownBy(() ->
                mapper.writerWithView(Views.Summary.class)
                        .writeValue(file, Arrays.asList(
                                new WithViews("Apple", "A fruit", 10))))
                .isInstanceOf(SheetStreamWriteException.class);
    }

    // ----------------------------------------------------------------
    // @JsonTypeInfo + @JsonSubTypes
    // ----------------------------------------------------------------

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
    @JsonSubTypes({
            @JsonSubTypes.Type(value = Dog.class, name = "dog"),
            @JsonSubTypes.Type(value = Cat.class, name = "cat")
    })
    static class Animal {
        public String name;
        public Animal() {}
    }

    static class Dog extends Animal {
        public String breed;
        public Dog() {}
        public Dog(String name, String breed) {
            this.name = name;
            this.breed = breed;
        }
    }

    static class Cat extends Animal {
        public boolean indoor;
        public Cat() {}
        public Cat(String name, boolean indoor) {
            this.name = name;
            this.indoor = indoor;
        }
    }

    @DataGrid
    static class WithPolymorphic {
        public String owner;
        public Animal pet;

        public WithPolymorphic() {}
        public WithPolymorphic(String owner, Animal pet) {
            this.owner = owner;
            this.pet = pet;
        }
    }

    @Test
    void jsonTypeInfo_unionSchema() throws Exception {
        File file = tempFile("polymorphic.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new WithPolymorphic("Alice", new Dog("Rex", "Labrador")),
                new WithPolymorphic("Bob", new Cat("Whiskers", true))),
                WithPolymorphic.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("owner");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("type");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("pet/name");
            assertThat(header.getCell(3).getStringCellValue()).isEqualTo("pet/breed");
            assertThat(header.getCell(4).getStringCellValue()).isEqualTo("pet/indoor");

            Row dogRow = wb.getSheetAt(0).getRow(1);
            assertThat(dogRow.getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(dogRow.getCell(1).getStringCellValue()).isEqualTo("dog");
            assertThat(dogRow.getCell(2).getStringCellValue()).isEqualTo("Rex");
            assertThat(dogRow.getCell(3).getStringCellValue()).isEqualTo("Labrador");

            Row catRow = wb.getSheetAt(0).getRow(2);
            assertThat(catRow.getCell(0).getStringCellValue()).isEqualTo("Bob");
            assertThat(catRow.getCell(1).getStringCellValue()).isEqualTo("cat");
            assertThat(catRow.getCell(2).getStringCellValue()).isEqualTo("Whiskers");
            assertThat(catRow.getCell(4).getBooleanCellValue()).isTrue();
        }
    }

    // -- Multi-level polymorphism --

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "cargoType")
    @JsonSubTypes({
            @JsonSubTypes.Type(value = Flatbed.class, name = "flatbed"),
            @JsonSubTypes.Type(value = Tanker.class, name = "tanker")
    })
    static class Cargo {
        public double weight;
        public Cargo() {}
    }

    static class Flatbed extends Cargo {
        public int axles;
        public Flatbed() {}
        public Flatbed(double weight, int axles) {
            this.weight = weight;
            this.axles = axles;
        }
    }

    static class Tanker extends Cargo {
        public String liquid;
        public Tanker() {}
        public Tanker(double weight, String liquid) {
            this.weight = weight;
            this.liquid = liquid;
        }
    }

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "vehicleType")
    @JsonSubTypes({
            @JsonSubTypes.Type(value = Sedan.class, name = "sedan"),
            @JsonSubTypes.Type(value = Truck.class, name = "truck")
    })
    static class Vehicle {
        public String model;
        public Vehicle() {}
    }

    static class Sedan extends Vehicle {
        public int doors;
        public Sedan() {}
        public Sedan(String model, int doors) {
            this.model = model;
            this.doors = doors;
        }
    }

    static class Truck extends Vehicle {
        public Cargo cargo;
        public Truck() {}
        public Truck(String model, Cargo cargo) {
            this.model = model;
            this.cargo = cargo;
        }
    }

    @DataGrid
    static class Fleet {
        public String company;
        public Vehicle vehicle;

        public Fleet() {}
        public Fleet(String company, Vehicle vehicle) {
            this.company = company;
            this.vehicle = vehicle;
        }
    }

    @Test
    void jsonTypeInfo_multiLevel() throws Exception {
        File file = tempFile("multi-poly.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new Fleet("A", new Sedan("Camry", 4)),
                new Fleet("B", new Truck("F-150", new Flatbed(5000, 3)))),
                Fleet.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("company");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("vehicleType");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("vehicle/model");

            Row sedanRow = wb.getSheetAt(0).getRow(1);
            assertThat(sedanRow.getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(sedanRow.getCell(1).getStringCellValue()).isEqualTo("sedan");
            assertThat(sedanRow.getCell(2).getStringCellValue()).isEqualTo("Camry");

            Row truckRow = wb.getSheetAt(0).getRow(2);
            assertThat(truckRow.getCell(0).getStringCellValue()).isEqualTo("B");
            assertThat(truckRow.getCell(1).getStringCellValue()).isEqualTo("truck");
            assertThat(truckRow.getCell(2).getStringCellValue()).isEqualTo("F-150");
        }
    }

    // ----------------------------------------------------------------
    // Mix-in
    // ----------------------------------------------------------------

    static class ThirdPartyDto {
        public String code;
        public int amount;

        public ThirdPartyDto() {}
        public ThirdPartyDto(String code, int amount) {
            this.code = code;
            this.amount = amount;
        }
    }

    @DataGrid
    abstract static class ThirdPartyMixin {
        @JsonProperty("Code")
        String code;
        @JsonProperty("Amount")
        int amount;
    }

    @Test
    void mixin_appliesAnnotations() throws Exception {
        mapper.addMixIn(ThirdPartyDto.class, ThirdPartyMixin.class);

        File file = tempFile("mixin.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new ThirdPartyDto("ABC", 100)), ThirdPartyDto.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("Code");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("Amount");
        }

        List<ThirdPartyDto> read = mapper.readValues(file, ThirdPartyDto.class);
        assertThat(read.get(0).code).isEqualTo("ABC");
        assertThat(read.get(0).amount).isEqualTo(100);
    }

    // ----------------------------------------------------------------
    // Column matching is positional
    // ----------------------------------------------------------------

    @DataGrid
    static class WithReorderedColumns {
        public String name;
        public int quantity;

        public WithReorderedColumns() {}
    }

    @Test
    void columnMatching_isPositional_byDefault() throws Exception {
        File file = tempFile("reordered.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("quantity");
            header.createCell(1).setCellValue("name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(10);
            data.createCell(1).setCellValue("Apple");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        assertThatThrownBy(() -> mapper.readValues(file, WithReorderedColumns.class))
                .hasMessageContaining("quantity");
    }

    @Test
    void columnReordering_matchesByHeaderName() throws Exception {
        File file = tempFile("reorder-match.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("quantity");
            header.createCell(1).setCellValue("name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(10);
            data.createCell(1).setCellValue("Apple");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper reorderMapper = SpreadsheetMapper.builder()
                .columnReordering(true).build();
        List<WithReorderedColumns> read = reorderMapper.readValues(
                file, WithReorderedColumns.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).quantity).isEqualTo(10);
    }

    @Test
    void columnReordering_extraColumnsIgnored() throws Exception {
        File file = tempFile("reorder-extra.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("quantity");
            header.createCell(1).setCellValue("extra");
            header.createCell(2).setCellValue("name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(10);
            data.createCell(1).setCellValue("ignored");
            data.createCell(2).setCellValue("Apple");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper reorderMapper = SpreadsheetMapper.builder()
                .columnReordering(true).build();
        List<WithReorderedColumns> read = reorderMapper.readValues(
                file, WithReorderedColumns.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).quantity).isEqualTo(10);
    }

    @Test
    void columnReordering_fewerColumnsInFile() throws Exception {
        File file = tempFile("reorder-fewer.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("Apple");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper reorderMapper = SpreadsheetMapper.builder()
                .columnReordering(true).build();
        List<WithReorderedColumns> read = reorderMapper.readValues(
                file, WithReorderedColumns.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).quantity).isEqualTo(0);
    }

    // ----------------------------------------------------------------
    // @JsonAlias — with columnReordering
    // ----------------------------------------------------------------

    @DataGrid
    static class WithAlias {
        @JsonAlias({"Product Name", "product_name"})
        public String name;
        public int quantity;

        public WithAlias() {}
    }

    @Test
    void jsonAlias_matchesHeaderWithColumnReordering() throws Exception {
        File file = tempFile("alias-reorder.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("quantity");
            header.createCell(1).setCellValue("Product Name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(10);
            data.createCell(1).setCellValue("Apple");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper reorderMapper = SpreadsheetMapper.builder()
                .columnReordering(true).build();
        List<WithAlias> read = reorderMapper.readValues(file, WithAlias.class);
        assertThat(read.get(0).name).isEqualTo("Apple");
        assertThat(read.get(0).quantity).isEqualTo(10);
    }

    // ----------------------------------------------------------------
    // Limited: @JsonAnySetter / @JsonAnyGetter
    // ----------------------------------------------------------------

    @DataGrid
    static class WithAnyProperties {
        public String name;
        private final Map<String, Object> extra = new HashMap<>();

        public WithAnyProperties() {}
        public WithAnyProperties(String name) { this.name = name; }

        @JsonAnyGetter
        public Map<String, Object> getExtra() { return extra; }
        @JsonAnySetter
        public void setExtra(String key, Object value) { extra.put(key, value); }
    }

    @Test
    void jsonAnyGetter_dynamicPropertiesExcludedFromSchema() throws Exception {
        File file = tempFile("any-setter.xlsx");
        WithAnyProperties item = new WithAnyProperties("test");
        item.setExtra("dynamic1", "value1");

        mapper.writeValue(file, Arrays.asList(item), WithAnyProperties.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Row header = wb.getSheetAt(0).getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1)).isNull();
        }
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
