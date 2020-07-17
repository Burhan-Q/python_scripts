#Author- Autodesk, modified by Burhan Qaddoumi
#Description-Extract BOM information from active design and output to text file.

import adsk.core, adsk.fusion, traceback, os

def spacePadRight(value, length):
    pad = ''
    if type(value) is str:
        paddingLength = length - len(value) + 1
    else:
        paddingLength = length - value + 1
    while paddingLength > 0:
        pad += ' '
        paddingLength -= 1

    return str(value) + pad

def walkThrough(bom):
    mStr = ''
    for item in bom:
        mStr += spacePadRight(item['name'], 25) + str(spacePadRight(item['instances'], 15)) + str(item['volume']) + '\n'
    return mStr

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'Extract BOM'
        if not design:
            ui.messageBox('No active design', title)
            return

        # Get all occurrences in the root component of the active design
        root = design.rootComponent
        occs = root.allOccurrences
        
        # Gather information about each unique component
        bom = []
        for occ in occs:
            comp = occ.component
            jj = 0
            for bomI in bom:
                if bomI['component'] == comp:
                    # Increment the instance count of the existing row.
                    bomI['instances'] += 1
                    break
                jj += 1

            if jj == len(bom):
                # Gather any BOM worthy values from the component
                volume = 0
                bodies = comp.bRepBodies
                for bodyK in bodies:
                    if bodyK.isSolid:
                        volume += bodyK.volume
                
                # Add this component to the BOM
                bom.append({
                    'component': comp,
                    'name': comp.name,
                    'instances': 1,
                    'volume': volume
                })

        # Display and export BOM to text file
        title = spacePadRight('Name', 25) + spacePadRight('Instances', 15) + 'Volume\n'
        msg = title + '\n' + walkThrough(bom)
        
        ui.messageBox(msg, 'Bill Of Materials')

        # Additional code to output BOM to text file; saved to user's Deskop
        docname = app.activeDocument.name
        name = docname + '_' + 'BOM.txt'
        path = os.path.join(os.path.expanduser('~'), 'desktop', name)
        f = open(path, 'w')
        f.write(walkThrough(bom))
        f.close()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
