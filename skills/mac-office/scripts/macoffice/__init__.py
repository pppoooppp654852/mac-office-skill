"""Implementation package for ``scripts/office_runner.py``.

Modules, one responsibility each:

  constants      app names, staging location, alert handlers, hint texts, OOXML namespaces
  text           text/column helpers shared by every inspector
  staging        per-run staging directory inside the Office Group Container
  applescript    osascript transport (AppleScript + JXA), staged-document cleanup, dialog probes
  app_session    one call's hold on an Office app: lock, hidden launch, focus guard, idle quit
  openxml        headless zipfile + ElementTree inspectors for .docx/.xlsx/.pptx
  container      OLE2 / MIP container detection
  fallback       native-app AppleScript inspectors used for OLE2 containers
  inspect        ``--inspect`` dispatch
  execute        ``--app`` execution, assets, lazy write-back
  render         ``--render`` PDF export (cached by content hash) and rasterization
"""
