package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.util.Arrays;

/**
 * In-memory shared string store using SoA layout:
 * one char slab + offsets/lengths arrays, with open-addressing dedup index.
 */
final class InMemorySharedStringsStore implements SharedStringsStore {

    private static final int INITIAL_CHAR_CAPACITY = 4096;
    private static final float LOAD_FACTOR = 0.7f;

    private char[] _data;
    private int[] _offsets;
    private int[] _lengths;
    private int[] _hashes;
    private int _dataPos;
    private int _size;

    private int[] _indexBySlot;
    private int _mask;
    private int _resizeThreshold;

    InMemorySharedStringsStore() {
        this(256);
    }

    InMemorySharedStringsStore(final int initialCapacity) {
        final int cap = _tableSizeFor(Math.max(16, initialCapacity));
        _indexBySlot = new int[cap];
        Arrays.fill(_indexBySlot, -1);
        _mask = cap - 1;
        _resizeThreshold = (int) (cap * LOAD_FACTOR);

        final int valueCap = Math.max(16, initialCapacity);
        _offsets = new int[valueCap];
        _lengths = new int[valueCap];
        _hashes = new int[valueCap];
        _data = new char[Math.max(INITIAL_CHAR_CAPACITY, valueCap * 16)];
    }

    @Override
    public int getOrAdd(final String value) {
        final String s = value == null ? "" : value;
        final int hash = _mix(s.hashCode());
        while (true) {
            int slot = hash & _mask;
            while (true) {
                final int index = _indexBySlot[slot];
                if (index < 0) {
                    if (_size >= _resizeThreshold) {
                        _resize();
                        break;
                    }
                    return _add(slot, s, hash);
                }
                if (_hashes[index] == hash
                        && _lengths[index] == s.length()
                        && _equals(index, s)) {
                    return index;
                }
                slot = (slot + 1) & _mask;
            }
        }
    }

    @Override
    public int size() {
        return _size;
    }

    @Override
    public String get(final int index) {
        return new String(_data, _offsets[index], _lengths[index]);
    }

    @Override
    public int length(final int index) {
        return _lengths[index];
    }

    @Override
    public char charAt(final int index, final int offset) {
        return _data[_offsets[index] + offset];
    }

    @Override
    public void close() {
        // no-op
    }

    private int _add(final int slot, final String s, final int hash) {
        final int index = _size;
        final int len = s.length();
        _ensureValueCapacity(index + 1);
        _ensureDataCapacity(len);

        final int offset = _dataPos;
        s.getChars(0, len, _data, offset);
        _dataPos += len;

        _offsets[index] = offset;
        _lengths[index] = len;
        _hashes[index] = hash;
        _indexBySlot[slot] = index;
        _size = index + 1;
        return index;
    }

    private boolean _equals(final int index, final String other) {
        final int len = _lengths[index];
        final int offset = _offsets[index];
        for (int i = 0; i < len; i++) {
            if (_data[offset + i] != other.charAt(i)) {
                return false;
            }
        }
        return true;
    }

    private void _ensureValueCapacity(final int required) {
        if (required <= _offsets.length) {
            return;
        }
        int next = _offsets.length;
        while (next < required) {
            next <<= 1;
        }
        _offsets = Arrays.copyOf(_offsets, next);
        _lengths = Arrays.copyOf(_lengths, next);
        _hashes = Arrays.copyOf(_hashes, next);
    }

    private void _ensureDataCapacity(final int additionalChars) {
        final int required = _dataPos + additionalChars;
        if (required <= _data.length) {
            return;
        }
        int next = _data.length;
        while (next < required) {
            next <<= 1;
        }
        _data = Arrays.copyOf(_data, next);
    }

    private void _resize() {
        final int newCap = _indexBySlot.length << 1;
        final int[] newIndexBySlot = new int[newCap];
        Arrays.fill(newIndexBySlot, -1);
        final int newMask = newCap - 1;

        for (int i = 0; i < _size; i++) {
            int slot = _hashes[i] & newMask;
            while (newIndexBySlot[slot] >= 0) {
                slot = (slot + 1) & newMask;
            }
            newIndexBySlot[slot] = i;
        }

        _indexBySlot = newIndexBySlot;
        _mask = newMask;
        _resizeThreshold = (int) (newCap * LOAD_FACTOR);
    }

    private static int _mix(final int x) {
        return x ^ (x >>> 16);
    }

    private static int _tableSizeFor(final int cap) {
        int n = cap - 1;
        n |= n >>> 1;
        n |= n >>> 2;
        n |= n >>> 4;
        n |= n >>> 8;
        n |= n >>> 16;
        return (n < 0) ? 1 : n + 1;
    }
}
