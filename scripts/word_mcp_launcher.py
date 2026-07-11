import sys
import os
import inspect

# We need to make sure pythoncom and win32com are available.
try:
    import pythoncom
    import win32com.client
except ImportError:
    pass

# Import the original module
try:
    import word_document_server.core.word_com as word_com
except ImportError:
    # If not running on Windows or package not found, just run the original entrypoint directly
    try:
        from word_document_server.main import run_server
        run_server()
        sys.exit(0)
    except Exception as e:
        print(f"Error starting server: {e}", file=sys.stderr)
        sys.exit(1)

# Keep references to the original functions
original_get_word_app = word_com.get_word_app
original_find_open_document = word_com.find_open_document

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

def get_all_word_apps():
    apps = []
    try:
        pythoncom.CoInitialize()
        rot = pythoncom.GetRunningObjectTable()
        monikers = rot.EnumRunning()
        
        for moniker in monikers:
            try:
                ctx = pythoncom.CreateBindCtx(0)
                display_name = moniker.GetDisplayName(ctx, None)
                
                # Check if moniker looks like a Word document or application
                if (display_name.startswith("!") or 
                    display_name.lower().endswith(".docx") or 
                    display_name.lower().endswith(".doc") or 
                    display_name.lower().endswith(".docm") or
                    "word" in display_name.lower()):
                    
                    unk = rot.GetObject(moniker)
                    disp = win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
                    
                    if hasattr(disp, "Application"):
                        apps.append(disp.Application)
                    elif hasattr(disp, "Documents"):
                        apps.append(disp)
            except Exception:
                pass
    except Exception:
        pass
        
    # De-duplicate apps by Hwnd
    unique_apps = []
    seen_hwnds = set()
    for app in apps:
        try:
            hwnd = app.Hwnd
            if hwnd not in seen_hwnds:
                seen_hwnds.add(hwnd)
                unique_apps.append(app)
        except Exception:
            unique_apps.append(app)
            
    return unique_apps

def patched_get_word_app():
    filename = get_filename_from_stack()
    if filename:
        target = filename.lower().replace('/', '\\')
        apps = get_all_word_apps()
        for app in apps:
            try:
                for doc in app.Documents:
                    doc_name = doc.Name.lower()
                    doc_fullname = doc.FullName.lower().replace('/', '\\')
                    if target == doc_name or target == doc_fullname or doc_fullname.endswith(target) or target.endswith(doc_name):
                        return app
            except Exception:
                pass
                
        # Try direct GetObject if it's a file path
        if '\\' in filename or '/' in filename or ':' in filename:
            try:
                doc = win32com.client.GetObject(filename)
                if doc and hasattr(doc, "Application"):
                    return doc.Application
            except Exception:
                pass

    # Fallback to original
    try:
        return original_get_word_app()
    except Exception:
        return win32com.client.GetActiveObject("Word.Application")

def patched_find_open_document(app, filename=None):
    if not filename:
        try:
            return app.ActiveDocument
        except Exception:
            pass
            
    target = filename.lower().replace('/', '\\')
    
    # Try searching in the provided app instance first
    try:
        for doc in app.Documents:
            doc_name = doc.Name.lower()
            doc_fullname = doc.FullName.lower().replace('/', '\\')
            if target == doc_name or target == doc_fullname or doc_fullname.endswith(target) or target.endswith(doc_name):
                return doc
    except Exception:
        pass

    # Search across all instances in the ROT
    apps = get_all_word_apps()
    for other_app in apps:
        try:
            for doc in other_app.Documents:
                doc_name = doc.Name.lower()
                doc_fullname = doc.FullName.lower().replace('/', '\\')
                if target == doc_name or target == doc_fullname or doc_fullname.endswith(target) or target.endswith(doc_name):
                    return doc
        except Exception:
            pass
            
    # Try direct GetObject
    if '\\' in filename or '/' in filename or ':' in filename:
        try:
            doc = win32com.client.GetObject(filename)
            if doc:
                return doc
        except Exception:
            pass
            
    # Fallback to original
    try:
        return original_find_open_document(app, filename)
    except Exception:
        return None

# Apply the patches
word_com.get_word_app = patched_get_word_app
word_com.find_open_document = patched_find_open_document

# Run the server
if __name__ == '__main__':
    from word_document_server.main import run_server
    run_server()
