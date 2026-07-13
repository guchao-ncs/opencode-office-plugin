import sys
import os
import inspect

# We need to make sure pythoncom and win32com are available.
try:
    import pythoncom
    import win32com.client
except ImportError:
    pass

# Import the original module.
# Assuming a structure similar to the Word one: `powerpoint_document_server.core.powerpoint_com`.
# If the PowerPoint server package name or module name is finalized to something else,
# adjust these imports accordingly.
try:
    import powerpoint_document_server.core.powerpoint_com as powerpoint_com
except ImportError:
    # If not running on Windows or package not found, just run the original entrypoint directly
    try:
        from powerpoint_document_server.main import run_server
        run_server()
        sys.exit(0)
    except Exception as e:
        print(f"Error starting server: {e}", file=sys.stderr)
        sys.exit(1)

# Keep references to the original functions.
# Adjust these attributes if the function names in the core PowerPoint COM module differ.
original_get_powerpoint_app = getattr(powerpoint_com, "get_powerpoint_app", None)
original_find_presentation = getattr(powerpoint_com, "find_presentation", None)

def get_filename_from_stack():
    try:
        for frame_info in inspect.stack():
            frame = frame_info.frame
            # Check if the frame has a local variable named 'filename'
            if 'filename' in frame.f_locals:
                filename = frame.f_locals['filename']
                if filename:
                    return filename
    except Exception:
        pass
    return None

def get_all_powerpoint_apps():
    apps = []
    try:
        pythoncom.CoInitialize()
        rot = pythoncom.GetRunningObjectTable()
        monikers = rot.EnumRunning()
        
        for moniker in monikers:
            try:
                ctx = pythoncom.CreateBindCtx(0)
                display_name = moniker.GetDisplayName(ctx, None)
                
                # Check if moniker looks like a PowerPoint presentation or application
                if (display_name.startswith("!") or 
                    display_name.lower().endswith(".pptx") or 
                    display_name.lower().endswith(".ppt") or 
                    display_name.lower().endswith(".pptm") or
                    "powerpoint" in display_name.lower()):
                    
                    unk = rot.GetObject(moniker)
                    disp = win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
                    
                    if hasattr(disp, "Application"):
                        apps.append(disp.Application)
                    elif hasattr(disp, "Presentations"):
                        apps.append(disp)
            except Exception:
                pass
    except Exception:
        pass
        
    # De-duplicate apps by Hwnd or HWND
    unique_apps = []
    seen_hwnds = set()
    for app in apps:
        try:
            # PowerPoint Application object can expose HWND (or Hwnd)
            hwnd = getattr(app, "HWND", getattr(app, "Hwnd", None))
            if hwnd is not None:
                if hwnd not in seen_hwnds:
                    seen_hwnds.add(hwnd)
                    unique_apps.append(app)
            else:
                unique_apps.append(app)
        except Exception:
            unique_apps.append(app)
            
    return unique_apps

def patched_get_powerpoint_app():
    filename = get_filename_from_stack()
    if filename:
        target = filename.lower().replace('/', '\\')
        apps = get_all_powerpoint_apps()
        for app in apps:
            try:
                for pres in app.Presentations:
                    pres_name = pres.Name.lower()
                    pres_fullname = pres.FullName.lower().replace('/', '\\')
                    if target == pres_name or target == pres_fullname or pres_fullname.endswith(target) or target.endswith(pres_name):
                        return app
            except Exception:
                pass
                
        # Try direct GetObject if it's a file path
        if '\\' in filename or '/' in filename or ':' in filename:
            try:
                pres = win32com.client.GetObject(filename)
                if pres and hasattr(pres, "Application"):
                    return pres.Application
            except Exception:
                pass

    # Fallback to original
    try:
        if original_get_powerpoint_app:
            return original_get_powerpoint_app()
    except Exception:
        pass
    return win32com.client.GetActiveObject("PowerPoint.Application")

def patched_find_presentation(app, filename=None):
    if not filename:
        try:
            return app.ActivePresentation
        except Exception:
            pass
            
    target = filename.lower().replace('/', '\\')
    
    # Try searching in the provided app instance first
    try:
        for pres in app.Presentations:
            pres_name = pres.Name.lower()
            pres_fullname = pres.FullName.lower().replace('/', '\\')
            if target == pres_name or target == pres_fullname or pres_fullname.endswith(target) or target.endswith(pres_name):
                return pres
    except Exception:
        pass

    # Search across all instances in the ROT
    apps = get_all_powerpoint_apps()
    for other_app in apps:
        try:
            for pres in other_app.Presentations:
                pres_name = pres.Name.lower()
                pres_fullname = pres.FullName.lower().replace('/', '\\')
                if target == pres_name or target == pres_fullname or pres_fullname.endswith(target) or target.endswith(pres_name):
                    return pres
        except Exception:
            pass
            
    # Try direct GetObject
    if '\\' in filename or '/' in filename or ':' in filename:
        try:
            pres = win32com.client.GetObject(filename)
            if pres:
                return pres
        except Exception:
            pass
            
    # Fallback to original
    try:
        if original_find_presentation:
            return original_find_presentation(app, filename)
    except Exception:
        pass
    return None

# Apply the patches if the module was imported successfully
if 'powerpoint_com' in locals() or 'powerpoint_com' in globals():
    powerpoint_com.get_powerpoint_app = patched_get_powerpoint_app
    powerpoint_com.find_presentation = patched_find_presentation

# Run the server
if __name__ == '__main__':
    from powerpoint_document_server.main import run_server
    run_server()
