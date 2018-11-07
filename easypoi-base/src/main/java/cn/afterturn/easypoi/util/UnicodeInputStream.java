package cn.afterturn.easypoi.util;

/**
 * @author by jueyue on 18-11-7.
 */

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

public class UnicodeInputStream extends InputStream {
    public static final UnicodeInputStream.UnicodeBOM UTF8_BOM = new UnicodeInputStream.UnicodeBOM("UTF-8", new byte[]{-17, -69, -65});
    public static final UnicodeInputStream.UnicodeBOM UTF16LE_BOM = new UnicodeInputStream.UnicodeBOM("UTF-16LE", new byte[]{-1, -2});
    public static final UnicodeInputStream.UnicodeBOM UTF16BE_BOM = new UnicodeInputStream.UnicodeBOM("UTF-16BE", new byte[]{-2, -1});
    public static final UnicodeInputStream.UnicodeBOM UTF32LE_BOM = new UnicodeInputStream.UnicodeBOM("UTF-32LE", new byte[]{-1, -2, 0, 0});
    public static final UnicodeInputStream.UnicodeBOM UTF32BE_BOM = new UnicodeInputStream.UnicodeBOM("UTF-32BE", new byte[]{0, 0, -2, -1});
    private static final int MAX_BOM_SIZE = 4;
    private byte[] buf;
    private int pos;
    private final String encoding;
    private final boolean skipBOM;
    private final PushbackInputStream inputStream;

    public UnicodeInputStream(InputStream inputStream) throws IllegalStateException, IOException {
        this(inputStream, true);
    }

    public UnicodeInputStream(InputStream inputStream, boolean skipBOM) throws IllegalStateException, IOException {
        this.buf = new byte[MAX_BOM_SIZE];
        this.pos = 0;
        this.skipBOM = skipBOM;
        this.inputStream = new PushbackInputStream(inputStream, MAX_BOM_SIZE);

        try {
            this.encoding = this.readEncoding();
        } catch (IOException var5) {
            IllegalStateException ex = new IllegalStateException("Could not read BOM from Stream");
            throw ex;
        }
    }

    public boolean isSkipBOM() {
        return this.skipBOM;
    }

    public String getEncodingFromStream() {
        return this.encoding;
    }

    protected String readEncoding() throws IOException {
        this.pos = 0;
        UnicodeInputStream.UnicodeBOM encoding = null;
        if (this.readByte()) {
            switch(this.buf[0]) {
                case -17:
                    encoding = this.match(UTF8_BOM, (UnicodeInputStream.UnicodeBOM)null);
                    break;
                case -2:
                    encoding = this.match(UTF16BE_BOM, (UnicodeInputStream.UnicodeBOM)null);
                    break;
                case -1:
                    encoding = this.match(UTF16LE_BOM, (UnicodeInputStream.UnicodeBOM)null);
                    if (encoding != null) {
                        encoding = this.match(UTF32LE_BOM, encoding);
                    }
                    break;
                case 0:
                    encoding = this.match(UTF32BE_BOM, (UnicodeInputStream.UnicodeBOM)null);
                    break;
                default:
                    encoding = null;
            }
        }

        this.pushback(encoding);
        return encoding != null ? encoding.getEncoding() : null;
    }

    private final UnicodeInputStream.UnicodeBOM match(UnicodeInputStream.UnicodeBOM matchEncoding, UnicodeInputStream.UnicodeBOM noMatchEncoding) throws IOException {
        byte[] bom = matchEncoding.getBytes();

        for(int i = 0; i < bom.length; ++i) {
            if (this.pos <= i && !this.readByte()) {
                return noMatchEncoding;
            }

            if (bom[i] != this.buf[i]) {
                return noMatchEncoding;
            }
        }

        return matchEncoding;
    }

    private final boolean readByte() throws IOException {
        int res = this.inputStream.read();
        if (res == -1) {
            return false;
        } else if (this.pos >= this.buf.length) {
            throw new IOException("BOM read error");
        } else {
            this.buf[this.pos++] = (byte)res;
            return true;
        }
    }

    private final void pushback(UnicodeInputStream.UnicodeBOM matchBOM) throws IOException {
        int count = this.pos;
        int start = 0;
        if (matchBOM != null && this.skipBOM) {
            start = matchBOM.getBytes().length;
            count = this.pos - start;
            if (count < 0) {
                throw new IllegalStateException("Match has more bytes than available!");
            }
        }

        this.inputStream.unread(this.buf, start, count);
    }

    @Override
    public void close() throws IOException {
        this.inputStream.close();
    }

    @Override
    public int available() throws IOException {
        return this.inputStream.available();
    }

    @Override
    public void mark(int readlimit) {
        this.inputStream.mark(readlimit);
    }

    @Override
    public boolean markSupported() {
        return this.inputStream.markSupported();
    }

    @Override
    public int read() throws IOException {
        return this.inputStream.read();
    }

    @Override
    public int read(byte[] b) throws IOException {
        return this.inputStream.read(b);
    }

    @Override
    public int read(byte[] b, int off, int len) throws IOException {
        return this.inputStream.read(b, off, len);
    }

    @Override
    public void reset() throws IOException {
        this.inputStream.reset();
    }

    @Override
    public long skip(long n) throws IOException {
        return this.inputStream.skip(n);
    }

    static final class UnicodeBOM {
        private final String encoding;
        private final byte[] bytes;

        private UnicodeBOM(String encoding, byte[] bytes) {
            this.encoding = encoding;
            this.bytes = bytes;
        }

        String getEncoding() {
            return this.encoding;
        }

        byte[] getBytes() {
            return this.bytes;
        }
    }
}