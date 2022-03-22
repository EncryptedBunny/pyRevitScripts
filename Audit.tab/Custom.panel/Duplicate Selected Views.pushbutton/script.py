"""Batch duplicates the selected views with or without detailing."""

from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

# Set component Metadata
__title__ = 'Duplicate Views'
__author__ = 'Sebastian Teo, pyRevit'
__helpurl__ = 'mailto:steo@hayball.com.au'

logger = script.get_logger()


def duplicableview(view):
    return view.CanViewBeDuplicated(DB.ViewDuplicateOption.Duplicate)


def duplicate_views(viewlist, with_detailing=True):
    with revit.Transaction('Duplicate selected views'):
        for el in viewlist:
            if with_detailing:
                dupop = DB.ViewDuplicateOption.WithDetailing
            else:
                dupop = DB.ViewDuplicateOption.Duplicate

            try:
                el.Duplicate(dupop)
            except Exception as duplerr:
                logger.error('Error duplicating view "{}" | {}'
                             .format(revit.query.get_name(el), duplerr))


selected_views = forms.select_views(filterfunc=duplicableview)

if selected_views:
    selected_option = \
        forms.CommandSwitchWindow.show(
            ['WITH Detailing',
             'WITHOUT Detailing'],
            message='Select duplication option:'
            )
    selected_no = \
        forms.ask_for_string("1","Enter Number of Duplicates","Duplicates")
    duplicates = int(selected_no)

    if selected_option:
        for i in range(duplicates):

            duplicate_views(
                selected_views,
                with_detailing=True if selected_option == 'WITH Detailing'
                else False)
