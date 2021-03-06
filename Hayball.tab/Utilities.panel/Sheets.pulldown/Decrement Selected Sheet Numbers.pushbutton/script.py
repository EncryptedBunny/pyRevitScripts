#pylint: disable=W0703,E0401,C0103,C0111
from pyrevit import coreutils
from pyrevit import revit
from pyrevit import forms
from pyrevit import script


__doc__ = 'Decreases the sheet number of the selected sheets by one. '\
          'The sheet name change will be printed if logging is set '\
          'to Verbose in pyRevit settings.'


selection = revit.get_selection()
logger = script.get_logger()


shift = -1

selected_sheets = forms.select_sheets(title='Select Sheets')
if not selected_sheets:
    script.exit()

sorted_sheet_list = sorted(selected_sheets, key=lambda x: x.SheetNumber)
if shift >= 0:
    sorted_sheet_list.reverse()
with revit.TransactionGroup('Shift Sheets'):
    for sheet in sorted_sheet_list:
        with revit.Transaction('Shift Single Sheet'):
            try:
                cur_sheet_num = sheet.SheetNumber
                sheet_num_param = sheet.LookupParameter('Sheet Number')
                sheet_num_param.Set(coreutils.decrement_str(sheet.SheetNumber,
                                                            shift))
                new_sheet_num = sheet.SheetNumber
                logger.info('{} -> {}'.format(cur_sheet_num, new_sheet_num))
            except Exception as shift_err:
                logger.error(shift_err)

            revit.doc.Regenerate()
