import logging, os, inspect
from PyQt5 import QtWidgets, QtCore

# Helper function to get current file name and line number
def get_file_and_line():
    frame = inspect.currentframe().f_back
    return f"{os.path.basename(frame.f_code.co_filename)}:{frame.f_lineno}"

class ProgressDialog:
    """
    A utility class to manage progress dialogs across the application.
    This provides consistent progress dialog behavior and appearance.
    """
    
    def __init__(self, parent=None):
        """
        Initialize the ProgressDialog handler.
        
        Args:
            parent: The parent widget for the progress dialog
        """
        self.parent = parent
        self.dialog = None
    
    def create(self, title="Tiến trình", initial_text="Đang xử lý..."):
        """
        Creates and shows a progress dialog.
        
        Args:
            title: The window title for the dialog
            initial_text: The initial status text to display
            
        Returns:
            Self for method chaining
        """
        try:
            self.dialog = QtWidgets.QProgressDialog(initial_text, None, 0, 100, self.parent)
            self.dialog.setWindowTitle(title)
            self.dialog.setWindowModality(QtCore.Qt.WindowModal)
            self.dialog.setCancelButton(None)  # No cancel button
            self.dialog.setMinimumDuration(0)  # Show immediately
            self.dialog.setValue(10)  # Start at 10%
            self.dialog.show()
            QtCore.QCoreApplication.processEvents()
            return self
        except Exception as e:
            logging.error(f"Error creating progress dialog at {get_file_and_line()}: {e}")
            return self
    
    def update(self, value, text=None):
        """
        Updates the progress dialog with new value and optionally new text.
        
        Args:
            value: The new progress value (0-100)
            text: Optional new text to display
            
        Returns:
            Self for method chaining
        """
        if self.dialog:
            try:
                if text:
                    self.dialog.setLabelText(text)
                self.dialog.setValue(value)
                QtCore.QCoreApplication.processEvents()
            except Exception as e:
                logging.error(f"Error updating progress dialog at {get_file_and_line()}: {e}")
        return self
    
    def close(self):
        """
        Safely closes the progress dialog.
        
        Returns:
            Self for method chaining
        """
        if self.dialog:
            try:
                self.dialog.setValue(100)
                QtCore.QCoreApplication.processEvents()
                self.dialog.close()
                self.dialog = None
            except Exception as e:
                logging.error(f"Error closing progress dialog at {get_file_and_line()}: {e}")
        return self
