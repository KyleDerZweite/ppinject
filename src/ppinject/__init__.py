from importlib.metadata import PackageNotFoundError, version

from ppinject.highlevel import RenderReport, render_template_slide

try:
	__version__ = version("ppinject")
except PackageNotFoundError:
	__version__ = "0.1.0b1"

__all__ = [
	"RenderReport",
	"render_template_slide",
	"__version__",
]
