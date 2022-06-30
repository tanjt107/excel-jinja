import numpy as np
import os
import jinja2
import pandas as pd
import shutil
from dataclasses import dataclass
from jinja2 import Environment, FileSystemLoader, select_autoescape


class Workbook:
    """A class to represent an Excel Workbook."""

    def __init__(self, root_dir: str):
        """
        Attributes:
            root_dir: a directory that will be the root directory of the archive.
                      For example, we typically chdir into root_dir before creating the archive.
        """
        self.root_dir = root_dir

    def to_xlsx(self, base_name: str) -> None:
        """
        Zip xmls to an Excel file.
        Parameters:
            base_name: name of the file to create, including the path,
                       minus any format-specific extension.
        """
        shutil.make_archive(base_name, "zip", self.root_dir)
        if os.path.exists(f"{base_name}.xlsx"):
            os.remove(f"{base_name}.xlsx")
        os.rename(f"{base_name}.zip", f"{base_name}.xlsx")
        shutil.rmtree(self.root_dir)

    def copy(self, dst: str = "temp"):
        """
        Make a copy of the workbook's xml.
        Parameters:
            dst: A string representing the path of the destination file or directory.
        Return:
            Workbook object.
        """
        os.makedirs(dst, exist_ok=True)
        for dir in os.listdir(self.root_dir):
            src = os.path.join(self.root_dir, dir)
            if os.path.isdir(src):
                shutil.copytree(src, os.path.join(dst, dir))
            else:
                shutil.copy2(src, os.path.join(dst, dir))
        return Workbook(dst)

    def stream_and_dump_template(self, **kwargs) -> None:
        """Render and dump the complete template into the workbook xml."""
        sheet = Worksheet._stream_template("workbook.j2", "templates/jinja", **kwargs)
        sheet.dump(f"{self.root_dir}/xl/workbook.xml")


@dataclass
class Worksheet:
    """
    A class to represent an Excel worksheet.
    Attributes:
        sheet: Sheet Number.
    """

    sheet: int

    @staticmethod
    def _stream_template(
        template_name: str, loader: str, **kwargs
    ) -> jinja2.environment.TemplateStream:
        """
        Parameters:
            loader: The template loader for this environment.
            template_name: Name of the template to load.
            kwargs
        """
        env = Environment(
            loader=FileSystemLoader(loader), autoescape=select_autoescape()
        )
        template = env.get_template(template_name)
        return template.stream(**kwargs)

    def stream_and_dump_template(self, **kwargs):
        """Render and dump the complete template into the sheet xml."""
        sheet = self._stream_template(
            f"sheet{self.sheet}.j2", "templates/jinja", **kwargs
        )
        sheet.dump(f"temp/xl/worksheets/sheet{self.sheet}.xml")
