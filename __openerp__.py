{
    'name': 'Production CDT',
    'version': '1.1',
    'category': 'Human Resources',
    'complexity': "easy",
    'description': """

       """,
    'author': 'Yassine',
     'images': [
    ],
    'depends': ['hr','hr_paie','production','reporting'],
    'update_xml': ['production_cdt_sequence.xml',
                   'production_cdt_data.xml',
                   'security/hr_security.xml',
                   'security/ir.model.access.csv',
                   'production_cdt_view.xml',
                   'production_cdt_report.xml',
                   'production_cdt_menu.xml',
                   'production_cdt_board_view.xml',
                   ],
    'demo_xml': [],
    'test': [],
    'installable': True,
    'auto_install': False,
    #'certificate': '0063495605613',
}
# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
