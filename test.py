import CQCSniffer

cs = CQCSniffer.CQCSniffer('https://nww.cqc.nxp.com/CQC/', 'nxf44756', 'China#0303')
print(cs.checkActive())
print(cs.closeEvent('526325A', 'CQPR', 'NXF44756', 'CCE', 'Closed'))
print(cs.createEvent('526325A', 'CQPR', 'NXF44756', 'NXF44756', 'CCE', 'CCE'))