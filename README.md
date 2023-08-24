# dxf_electrical_panel

This project takes as an input a xlsx description of a electrical part and creates a dxf and svg shape.
For example:
………
………
………
Why to represent as rectangles:
The information you need is mainly the ports and the type of the part. So a simple rectangle with information such as the part type(ex circuit breaker,power supply,motor),part model,part comment,part label(unique for each part) could be sufficient to describe a schematic panel and most importantly the connections between the parts. Also trying to find the apropiete schematic could be time consuming given that for an part could be multiple schematics. Example encoders,inductive sensors 
Also when adding a part usually there is no need to add all the available ports of the parts,but only the used ones.So there is the option for each part there are Groups to represent the ports and their position. 
