package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.HeaderLayoutVisitor;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Applies header-cell comments (from {@code @DataColumn(comment)} and
 * {@code @DataColumnGroup(comment)}) to a POI {@link Sheet} for the given
 * {@link SpreadsheetSchema}. Internal write-time helper — invoked by the
 * sheet writers.
 */
public final class HeaderComments {

    private static final int COMMENT_BOX_WIDTH_COLS = 2;
    private static final int COMMENT_BOX_HEIGHT_ROWS = 3;

    private HeaderComments() {}

    public static void apply(final Sheet sheet, final SpreadsheetSchema schema) {
        if (!schema.usesHeader()) return;
        final CreationHelper factory = sheet.getWorkbook().getCreationHelper();
        schema.forEachHeaderCell(new HeaderLayoutVisitor() {
            private Drawing<?> drawing;

            @Override
            public void visitColumnHeader(final int row, final int col, final Column column) {
                _addCommentIfPresent(row, col, column.getValue().getComment());
            }

            @Override
            public void visitGroupCell(final int row, final int firstCol, final int lastCol,
                                       final DataColumnGroup.Value group) {
                _addCommentIfPresent(row, firstCol, group.getComment());
            }

            private void _addCommentIfPresent(final int row, final int col, final String text) {
                if (text.isEmpty()) return;
                if (drawing == null) drawing = sheet.createDrawingPatriarch();
                final ClientAnchor anchor = factory.createClientAnchor();
                anchor.setCol1(col);
                anchor.setRow1(row);
                anchor.setCol2(col + COMMENT_BOX_WIDTH_COLS);
                anchor.setRow2(row + COMMENT_BOX_HEIGHT_ROWS);
                final Comment c = drawing.createCellComment(anchor);
                c.setString(factory.createRichTextString(text));
                c.setAddress(row, col);
            }
        });
    }
}
