package support.fixture;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.FieldNameConstants;

import java.util.Arrays;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
@FieldNameConstants
@DataGrid
public class Entry {

    public static final Entry VALUE = new Entry(1, 2);
    public static final List<Entry> VALUES = Arrays.asList(new Entry(1, 2), new Entry(3, 4));
    public static final Entry[] VALUES_ARRAY = new Entry[]{new Entry(1, 2), new Entry(3, 4)};

    int a;
    int b;
}
