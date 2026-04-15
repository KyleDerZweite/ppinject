from importlib.metadata import PackageNotFoundError, version

from ppinject.highlevel import render_template_slide
from ppinject.injector import RenderReport

try:
	__version__ = version("ppinject")
except PackageNotFoundError:
	__version__ = "0.1.0b2"

__all__ = [
	"RenderReport",
	"render_template_slide",
	"__version__",
]
