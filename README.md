# ASEN - Analyzing Social or Economic Networks
A program for automatically calculating various sets of indexes/methods of network analysis. The aim of this program is applying Network Analysis methods for non-coders.  

The description/documentation of indicators/methods and its limits of applicability can be found at https://www.luciobiggiero.com/  A more detailed introduction to some indexes - especially the less known - can be found in the Methodological Appendix of the book "Inter-firm Networks", written by Lucio Biggiero & Robert Magnuszewski for the Springer Series on Relational Economics and Organizational Governance. It can be downloaded for free (by clicking on Back Matter) from https://link.springer.com/book/10.1007/978-3-031-17389-9

![screenshot](https://github.com/mbiggiero/ASEN/blob/main/screenshot.png?raw=true) 


Choose a graph (.xls(x) matrix/edgelist, UCINET's DL or pickled NetworkX graph), select a group of indicators and click "Analyze". 
Once a network is submitted to the program, it is automatically diagnosed whether directed/undirected, connected/disconnected, binary/weighted, including/excluding self-links.
Output files (.txt and .xlsx) will appear in the folder after the analysis is completed.

Note: most of implemented indicators are imported (and sometimes modified) from NetworkX, while the rest are written from scratch, based on network analysis concepts and methods. Because many indicators have limits of applicability, before choosing please check the complete reference list (and preferably even the code).  

Note for coders: this Python script was written with modularity in mind, so adding/modifying indicators/groups for your own use cases should be pretty straightforward.
