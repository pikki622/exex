from cleo import Application
from cleo import Command

from extract import Extractor


class ExtractCommand(Command):
    """
    Extract data from Excel document.

    extract
        {file : Source file path}
        {range? : Range}
        {format? : Format (json, jsonl, xml)}
        {skip_rows? : Number of rows to skip}
        {skip_cols? : Number of cols to skip}
    """

    def handle(self):
        excel_file_path = self.argument("file")

        if not excel_file_path:
            self.info("FAIL")
            return

        if self.argument("format"):
            pass

        ext = Extractor(excel_file_path)
        values = ext.all()

        print(values)


application = Application(name="exex", version="0.1.0")
application.add(ExtractCommand())

if __name__ == "__main__":
    application.run()
