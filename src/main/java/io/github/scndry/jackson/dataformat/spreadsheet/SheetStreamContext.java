package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.JsonStreamContext;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.apache.poi.ss.util.CellAddress;

public abstract class SheetStreamContext extends JsonStreamContext {

    protected static final int INITIAL_INDEX = -1;
    protected final SpreadsheetSchema _schema;
    protected int _size;

    protected SheetStreamContext(final int type, final SpreadsheetSchema schema) {
        super(type, INITIAL_INDEX);
        _schema = schema;
    }

    public static SheetStreamContext createRootContext(final SpreadsheetSchema schema) {
        return new RootContext(schema);
    }

    @Override
    public SheetStreamContext getParent() {
        return null;
    }

    @SuppressWarnings("java:S1172") // Unused method parameters should be removed
    public SheetStreamContext getParent(final Matcher matcher) {
        return null;
    }

    @Override
    public String getCurrentName() {
        return null;
    }

    public void setCurrentName(final String name) {
    }

    public SheetStreamContext clearAndGetParent() {
        return getParent();
    }

    public SheetStreamContext createChildArrayContext() {
        return createChildArrayContext(-1);
    }

    public SheetStreamContext createChildArrayContext(final int size) {
        return new ArrayContext(this, size);
    }

    public SheetStreamContext createChildObjectContext() {
        return new ObjectContext(this);
    }

    public int size() {
        return _size;
    }

    public abstract int getRow();

    public abstract int getColumn();

    public abstract void writeValue();

    public CellAddress currentReference() {
        return new CellAddress(getRow(), getColumn());
    }

    public abstract ColumnPointer currentPointer();

    public ColumnPointer relativePointer(final ColumnPointer other) {
        return getParent().currentPointer().relativize(other);
    }

    interface StepAware {

        int DEFAULT_STEP = 1;

        int getStep();

        void setStep(int step);
    }

    interface Matcher {
        boolean matches(SheetStreamContext context);
    }

    static final class RootContext extends SheetStreamContext implements StepAware {

        private int _step = DEFAULT_STEP;

        RootContext(final SpreadsheetSchema schema) {
            super(TYPE_ROOT, schema);
        }

        @Override
        public ColumnPointer currentPointer() {
            return ColumnPointer.empty();
        }

        @Override
        public int getRow() {
            return _schema.getDataRow() + _index;
        }

        @Override
        public int getColumn() {
            return _schema.getOriginColumn() + _index;
        }

        @Override
        public void writeValue() {
            _index += _step;
            _step = DEFAULT_STEP;
            _size = _index + 1;
        }

        @Override
        public int getStep() {
            return _step;
        }

        @Override
        public void setStep(final int step) {
            _step = step;
        }
    }

    abstract static class ChildContext extends SheetStreamContext {

        protected final SheetStreamContext _parent;

        ChildContext(final int type, final SheetStreamContext parent) {
            super(type, parent._schema);
            _parent = parent;
        }

        @Override
        public SheetStreamContext getParent() {
            return _parent;
        }

        @Override
        public SheetStreamContext getParent(final Matcher matcher) {
            return matcher.matches(_parent) ? _parent : _parent.getParent(matcher);
        }

        @Override
        public SheetStreamContext clearAndGetParent() {
            _parent._size += _size - 1;
            return super.clearAndGetParent();
        }
    }

    static final class ArrayContext extends ChildContext implements StepAware {

        private int _step = DEFAULT_STEP;

        ArrayContext(final SheetStreamContext parent, final int size) {
            super(TYPE_ARRAY, parent);
            _size = size;
        }

        @Override
        public ColumnPointer currentPointer() {
            if (_parent.inRoot()) return _parent.currentPointer();
            return _parent.currentPointer().resolveArray();
        }

        @Override
        public int getRow() {
            return _parent.getRow() + _index;
        }

        @Override
        public int getColumn() {
            return _parent.getColumn();
        }

        @Override
        public void writeValue() {
            _index += _step;
            _step = DEFAULT_STEP;
            _size = _index + 1;
        }

        @Override
        public SheetStreamContext clearAndGetParent() {
            final SheetStreamContext parent = super.getParent(StepAware.class::isInstance);
            if (parent != null) {
                ((StepAware) parent).setStep(_step + getRow() - parent.getRow());
            }
            return super.clearAndGetParent();
        }

        @Override
        public int getStep() {
            return _step;
        }

        @Override
        public void setStep(final int step) {
            _step = step;
        }
    }

    static final class ObjectContext extends ChildContext {

        private String _name;

        ObjectContext(final SheetStreamContext parent) {
            super(TYPE_OBJECT, parent);
        }

        @Override
        public String getCurrentName() {
            return _name;
        }

        @Override
        public void setCurrentName(final String name) {
            _name = name;
        }

        @Override
        public ColumnPointer currentPointer() {
            return _parent.currentPointer().resolve(_name);
        }

        @Override
        public int getRow() {
            return _parent.getRow();
        }

        @Override
        public int getColumn() {
            return hasCurrentIndex() ? _index : _parent.getColumn();
        }

        @Override
        public void writeValue() {
            _index = _schema.columnIndexOf(currentPointer());
            if (_size == 0) _size = 1;
        }
    }
}
