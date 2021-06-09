package op.tools.docx2wiki;

public class UploadBmpInfo {
    private String _FileName;
    private String _FileFullName;
    private String _Description = "";
    private String _Source = "";
    private String _Date = "";
    private String _Author = "";
    private String _Other_versions = "";
    private byte[] _BmpData;


    public String get_FileName() {
        return _FileName;
    }

    public void set_FileName(String _FileName) {
        this._FileName = _FileName;
    }

    public String get_FileFullName() {
        return _FileFullName;
    }

    public void set_FileFullName(String _FileFullName) {
        this._FileFullName = _FileFullName;
    }

    public String get_Description() {
        return _Description;
    }

    public void set_Description(String _Description) {
        this._Description = _Description;
    }

    public String get_Source() {
        return _Source;
    }

    public void set_Source(String _Source) {
        this._Source = _Source;
    }

    public String get_Date() {
        return _Date;
    }

    public void set_Date(String _Date) {
        this._Date = _Date;
    }

    public String get_Author() {
        return _Author;
    }

    public void set_Author(String _Author) {
        this._Author = _Author;
    }

    public String get_Other_versions() {
        return _Other_versions;
    }

    public void set_Other_versions(String _Other_versions) {
        this._Other_versions = _Other_versions;
    }

    public byte[] get_BmpData() {
        return _BmpData;
    }

    public void set_BmpData(byte[] _BmpData) {
        this._BmpData = _BmpData;
    }
}
