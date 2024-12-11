from PyQt6.QtWidgets import QWidget, QVBoxLayout
from ..utils.gui_components import FileDropArea
from ..utils.file_utils import process_file
from ..utils.date_utils import get_amazon_week

class BaseTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)

    def init_ui(self):
        raise NotImplementedError("Subclasses must implement init_ui method")

    def process(self):
        raise NotImplementedError("Subclasses must implement process method")

    def restart(self):
        raise NotImplementedError("Subclasses must implement restart method")