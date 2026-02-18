# -*- coding: utf-8 -*-
{
    'name': 'OM Bank Statement Import Custom',
    'version': '16.0.1.0.0',
    'category': 'Accounting',
    'summary': 'Import Bank Statement from CSV/XLSX',
    'description': """
        Custom module to import bank statements from CSV/XLSX files.
        Aggregates all transactions from a single file into one Bank Statement.
    """,
    'author': 'Antigravity',
    'depends': ['account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/bank_statement_import_view.xml',
        'views/account_journal_view.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
