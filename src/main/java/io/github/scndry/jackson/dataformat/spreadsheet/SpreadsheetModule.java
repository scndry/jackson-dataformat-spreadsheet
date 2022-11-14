package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.Version;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.DataGridBeanDeserializer;

final class SpreadsheetModule extends com.fasterxml.jackson.databind.Module {

    public static final String MODULE_NAME = "spreadsheet";

    @Override
    public String getModuleName() {
        return MODULE_NAME;
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    @Override
    public void setupModule(final SetupContext context) {
        context.addBeanDeserializerModifier(new DataGridBeanDeserializer.Modifier());
    }
}
