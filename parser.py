# -*- coding: utf-8 -*-
##############################################################################
#
# Copyright (c) 2008-2011 Alistek Ltd (http://www.alistek.com) All Rights Reserved.
#                    General contacts <info@alistek.com>
#
# WARNING: This program as such is intended to be used by professional
# programmers who take the whole responsability of assessing all potential
# consequences resulting from its eventual inadequacies and bugs
# End users who are looking for a ready-to-use solution with commercial
# garantees and support are strongly adviced to contract a Free Software
# Service Company
#
# This program is Free Software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 3
# of the License, or (at your option) any later version.
#
# This module is GPLv3 or newer and incompatible
# with OpenERP SA "AGPL + Private Use License"!
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
##############################################################################

import logging
logger = logging.getLogger('report_aeroo')

from openerp.report import report_sxw
from openerp.report.report_sxw import rml_parse
import time
import datetime

class Parser(rml_parse):
    def __init__(self, cr, uid, name, context):
        super(self.__class__, self).__init__(cr, uid, name, context)
        self.localcontext.update({
            'cr' : cr,
            'get_date' : self.get_date,
            'get_site' : self.get_site,
            'get_adress' : self.get_adress,
            'get_phone' : self.get_phone,
        })
    def get_date(self,cr,date):        
        return datetime.datetime.strptime(date, '%Y-%m-%d').strftime("%d/%m/%y")
    def get_site(self,cr,site):
        if site=='CASA' :
            rslt='TELEXCEL BPO'
        else :
            rslt='TELEXCEL ABIDJAN'
        return rslt
    def get_adress(self,cr,site):
        if site=='CASA' :
            rslt='5 Rue Abou Eddaboussi Casablanca'
        else :
            rslt='7 avenue Nogues, immeuble SCI Broadway (BSIC), 5ième étage 01 BP5754 abidjan 01'
        return rslt
    def get_phone(self,cr,site):
        if site=='CASA' :
            rslt='Tel : 00(212)522 36 60 31- Fax : 00(212)522 36 60 68'
        else :
            rslt='Tel : 00(225) 03 23 23 25'
        return rslt