
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.utils.clipboard import Clipboard

class TestClipboardAudit:
    """Audit clipboard edge cases."""

    def setup_method(self):
        self.sheet = Spreadsheet()
        self.clipboard = Clipboard(self.sheet)

    def test_copy_paste_overlap(self):
        """Audit: Copy A1:B2 to B2:C3 (overlapping)."""
        # A1=1, A2=2, B1=3, B2=4
        self.sheet.set_cell(0, 0, "1")
        self.sheet.set_cell(1, 0, "2")
        self.sheet.set_cell(0, 1, "3")
        self.sheet.set_cell(1, 1, "4")
        
        # Copy A1:B2
        self.clipboard.copy_range(0, 0, 1, 1)
        
        # Paste to B2
        self.clipboard.paste(1, 1)
        
        # B2 should be 1 (from old A1). 
        # C2 should be 3 (from old B1).
        # B3 should be 2 (from old A2).
        # C3 should be 4 (from old B2).
        assert self.sheet.get_value(1, 1) == 1.0
        assert self.sheet.get_value(1, 2) == 3.0
        assert self.sheet.get_value(2, 1) == 2.0
        assert self.sheet.get_value(2, 2) == 4.0

    def test_cut_paste_self_overlap(self):
        """Audit: Cut A1:B2 and Paste to A1 (no-op/restore)."""
        self.sheet.set_cell(0, 0, "1")
        self.clipboard.cut_range(0, 0, 1, 1)
        
        # Paste back to same spot
        self.clipboard.paste(0, 0)
        
        assert self.sheet.get_value(0, 0) == 1.0
        
    def test_paste_special_transpose_overlap(self):
        """Audit: Transpose paste on overlapping area."""
        # A1=1, A2=2. 
        self.sheet.set_cell(0, 0, "1")
        self.sheet.set_cell(1, 0, "2")
        
        self.clipboard.copy_range(0, 0, 1, 0) # Copy A1:A2 (2x1)
        
        # Transpose paste to A1. Should become A1:B1 (1x2).
        # A1 gets A1(1). B1 gets A2(2).
        # A2 is NOT cleared by copy, but might be overwritten? 
        # A2 is outside the new transposed destination range (A1:B1).
        self.clipboard.paste_special(0, 0, transpose=True)
        
        assert self.sheet.get_value(0, 0) == 1.0
        assert self.sheet.get_value(0, 1) == 2.0
        # A2 should remain 2 (unchanged by paste)
        assert self.sheet.get_value(1, 0) == 2.0

