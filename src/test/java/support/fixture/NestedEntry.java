package support.fixture;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@DataGrid
public class NestedEntry {

    int a;
    Inner inner;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static class Inner {
        int b;
    }
}
