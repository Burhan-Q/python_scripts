#Author: Burhan Qaddoumi
#Description: Export bodies/components as STL/STEP TODO: create GitHub commit, and add commit ID to version description in Fusion360
#Organization: Tetra Bio Distributed
#https://github.com/tetrabiodistributed
#https://www.tetrabio.org/

import adsk.core, adsk.fusion, adsk.cam, traceback
from pathlib import Path
from datetime import date

def run(context):
    ui = None
    
    #TODO create GitHub commit and save commit ID
    commit = ''
    
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        msg = f"Export complete, add files to {commit} on GitHub and submit Pull Request"
        ui.messageBox(msg, "Finished")
        
        #creates variable 'doc' from active Fusion 360 file as <class 'adsk.fusion.FusionDocument'>
        doc: adsk.core.Document = app.activeDocument
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        export = design.exportManager #export manager instance
        all_comps = design.allComponents
        
        #make directory for output
        home = Path.home()
        home_dirs = list(home.iterdir())
        s_date = str(date.today())
        new_dir = Path(str(home) + '/' + 'tetraBio' + '/' + s_date + '/' + doc.name)
        
        if not new_dir.exists() and new_dir not in home_dirs:
          try:
            new_dir.mkdir(parents=True)
          except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            raise ValueError(f"Could not create output directory at path {new_dir}")

        elif new_dir.exists() and (new_dir in home_dirs or new_dir not in home_dirs):
          new_dir.mkdir(parents=True,exist_ok=True) #will overwrite anything existing

        else:
          raise ValueError(f"Could not create output directory at path {new_dir}")

        #collect all components/bodies for export
        root = design.rootComponent
        all_ocur = root.allOccurrences #this returns an object containing information on the number of instances a body/component occurs

        #FUTURE file export naming via standard part numbering scheme

        #export all in collection as STL/STEP
        for occur in all_ocur:
          if not occur.fullPathName.startswith('Hardware') and occur._get_isLightBulbOn():
            #generate file name from component name and part number (assigned in Fusion360 Body/Component Properties)
            o_name = occur.component.partNumber + " - " + occur.component.name
            # STEP file output
            step_opt = export.createSTEPExportOptions(str(new_dir) + f"/{o_name}")
            export.execute(step_opt)

            # STL file output
            stl_opt = export.createSTLExportOptions(occur, (str(new_dir) + f"/{o_name}"))
            # stl_opt = export.createSTLExportOptions(str(new_dir) + f"/{o_name}") # old method, failed as of Fusion 2.0.15995
            export.execute(stl_opt)
        
        #save F360 file and add commit ID to description notes
        desc = f"Exported and uploaded as GitHub commit {commit}"
        
        #check if document is saved
        doc_saved = doc.isSaved
        
        if not doc_saved:
          #save with updated description, planned for future when GitHub commit value is available
          doc.save(desc)

        elif doc_saved:
          #toggle bodies folder visibility to enable new version save
          all_comps[0]._set_isBodiesFolderLightBulbOn(False)
          all_comps[0]._set_isBodiesFolderLightBulbOn(True)
          doc.save(desc)
          
          msg = f"Export complete, add files to {commit} on GitHub and submit Pull Request"
          ui.messageBox(msg, "Finished")
          
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

#Troubleshooting and/or notes:

#doc.dataFile.description #returns Version save description/notes
#NOTE: appears that after saving *sometimes* this information doesn't update
