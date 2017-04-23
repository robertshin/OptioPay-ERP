# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from frappe import _

def get_data():
	return [
		{
			"module_name": "Parties",
			"color": "#3498db",
			"icon": "octicon octicon-file-directory",
			"type": "module"
		},
		{
			"module_name": "Items",
			"color": "#f39c12",
			"icon": "octicon octicon-package",
			"type": "module"
		},
		{
			"module_name": "Transactions",
			"color": "#c0392b",
			"icon": "octicon octicon-briefcase",
			"type": "module"
		},
		{
			"module_name": "Invoices",
			"color": "#3498db",
			"icon": "octicon octicon-repo",
			"type": "module"
		},
	]