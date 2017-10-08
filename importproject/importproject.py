# coding=utf8

import mmap
import xlrd
import os

from genshi.builder import tag
from trac.core import *
from trac.ticket import Ticket, IPermissionRequestor
from trac.util.datefmt import FixedOffset
from trac.web.api import IRequestHandler
from trac.wiki.macros import WikiMacroBase
from trac.util.translation import _
from datetime import datetime


class ImportProject(Component):

    implements(IRequestHandler, IPermissionRequestor)

    def get_permission_actions(self):
        return ['IMPORT_PROJECT']

    def match_request(self, req):
        return req.method == 'POST' and req.path_info.startswith('/importproject')

    def process_request(self, req):

        assert req.method == 'POST'
        req.perm.assert_permission('IMPORT_PROJECT')

        upload = req.args.getfirst('attachment')
        if not hasattr(upload, 'filename') or not upload.filename:
            raise TracError(_("No file uploaded"))
        if hasattr(upload.file, 'fileno'):
            size = os.fstat(upload.file.fileno())[6]
        else:
            upload.file.seek(0, 2)
            size = upload.file.tell()
            upload.file.seek(0)
        if size == 0:
            raise TracError(_("Can't upload empty file"))

        book = xlrd.open_workbook(file_contents=mmap.mmap(upload.file.fileno(), 0, access=mmap.ACCESS_READ))

        sheet = book.sheet_by_name('import')

        ticket_main = Ticket(self.env)
        ticket_main['type'] = u'生产设计单'
        ticket_main['summary'] = sheet.cell(1, 3).value
        ticket_main['reporter'] = req.session.sid
        ticket_main['status'] = u'PJ01_可研'
        ticket_main['owner'] = self._get_user(sheet.cell(18, 3).value)
        ticket_main['project_short_name'] = self._get_str(sheet.cell(2, 3).value)
        ticket_main['description'] = self._get_str(sheet.cell(3, 3).value)
        ticket_main['project_no'] = self._get_str(sheet.cell(4, 3).value)
        ticket_main['project_feasibility_study_value'] = self._get_str(sheet.cell(5, 3).value)
        ticket_main['project_primary_design_value'] = self._get_str(sheet.cell(6, 3).value)
        ticket_main['project_detail_design_value'] = self._get_str(sheet.cell(7, 3).value)
        ticket_main['project_feasibility_study_value_electrical_ratio'] = self._get_str(sheet.cell(8, 3).value)
        ticket_main['project_feasibility_study_value_structure_ratio'] = self._get_str(sheet.cell(9, 3).value)
        ticket_main['project_primary_design_value_electrical_ratio'] = self._get_str(sheet.cell(10, 3).value)
        ticket_main['project_primary_design_value_structure_ratio'] = self._get_str(sheet.cell(11, 3).value)
        ticket_main['project_feasibility_study_value_electrical_ratio'] = self._get_str(sheet.cell(12, 3).value)
        ticket_main['project_feasibility_study_value_structure_ratio'] = self._get_str(sheet.cell(13, 3).value)
        ticket_main['project_electrical_drafter_ratio'] = self._get_str(sheet.cell(14, 3).value)
        ticket_main['project_electrical_reviewer_ratio'] = self._get_str(sheet.cell(15, 3).value)
        ticket_main['project_structure_drafter_ratio'] = self._get_str(sheet.cell(16, 3).value)
        ticket_main['project_structure_reviewer_ratio'] = self._get_str(sheet.cell(17, 3).value)
        ticket_main['project_manager'] = self._get_str(sheet.cell(18, 3).value)
        ticket_main['project_member_electrical_drafter'] = self._get_str(sheet.cell(19, 3).value)
        ticket_main['project_member_electrical_reviewer'] = self._get_str(sheet.cell(20, 3).value)
        ticket_main['project_member_structure_drafter'] = self._get_str(sheet.cell(21, 3).value)
        ticket_main['project_member_structure_reviewer'] = self._get_str(sheet.cell(22, 3).value)
        ticket_main['activity_start_date'] = self._get_date(sheet.cell(23, 3).value, book)
        ticket_main['activity_finish_date'] = self._get_date(sheet.cell(24, 3).value, book)
        ticket_main['project_location'] = self._get_str(sheet.cell(25, 3).value)
        ticket_main['project_type'] = self._get_str(sheet.cell(26, 3).value)
        ticket_main.insert()

        act_row = 28
        act_ids = []
        while act_row < sheet.nrows and sheet.cell(act_row, 1).value != '' \
                and sheet.cell(act_row, 1).value.upper() != 'N' and sheet.cell(act_row, 1).value.upper() != 'FALSE':
            ticket_act = Ticket(self.env)
            ticket_act['type'] = u'活动'
            ticket_act['summary'] = self._get_str(sheet.cell(act_row, 0).value)
            ticket_act['reporter'] = req.session.sid
            ticket_act['status'] = u'AC01_新活动'
            ticket_act['owner'] = None
            ticket_act['component'] = self._get_str(sheet.cell(act_row, 2).value)
            ticket_act['activity_type'] = self._get_str(sheet.cell(act_row, 3).value)
            ticket_act['activity_owner'] = self._get_str(sheet.cell(act_row, 4).value)
            ticket_act['activity_reviewer'] = self._get_str(sheet.cell(act_row, 5).value)
            ticket_act['activity_earn_value'] = self._get_str(sheet.cell(act_row, 6).value)
            ticket_act['activity_start_date'] = self._get_date(sheet.cell(act_row, 7).value, book)
            ticket_act['activity_finish_date'] = self._get_date(sheet.cell(act_row, 8).value, book)
            ticket_act.insert()
            act_ids.append(str(ticket_act.id))
            act_row += 1

        ticket_main['activity_project_relation_b'] = ','.join(act_ids)
        ticket_main.save_changes(req.session.sid, comment=u'导入设计生产单')

        req.redirect(req.href('ticket', ticket_main.id))

    def _get_user(self, user):

        users = self.env.get_known_users()

        for uid, uname, uemail in users:
            if uid == user or uname == user:
                return uid

        return None

    def _get_date(self, t, book):
        try:
            time = list(xlrd.xldate_as_tuple(t, book.datemode))
            time.append(0)
            time.append(FixedOffset(0, 'UTC'))
            result = datetime(*time)
            return result
        except:
            return None

    def _get_str(self, v):
        if v is None:
            return ''
        else:
            return unicode(v)

class ImportProjectMacro(WikiMacroBase):

    def expand_macro(self, formatter, name, content):

        return tag.form(tag.input(type='file', name='attachment'),
                        tag.input(type='hidden', name='__FORM_TOKEN', value=formatter.req.form_token),
                        tag.input(type='submit'),action='/importproject', method='POST', enctype="multipart/form-data")