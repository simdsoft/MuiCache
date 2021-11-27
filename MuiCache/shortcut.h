// Copyright (c) 2012 The Chromium Authors. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#ifndef BASE_WIN_SHORTCUT_H_
#define BASE_WIN_SHORTCUT_H_

#include <windows.h>
#include <commctrl.h>
#include <stdint.h>
#include <objbase.h>
#include <propkey.h>
#include <shellapi.h>
#include <shlobj.h>
#include <propvarutil.h>

namespace base
{
namespace win
{

// Properties for shortcuts. Properties set will be applied to the shortcut on
// creation/update, others will be ignored.
// Callers are encouraged to use the setters provided which take care of
// setting |options| as desired.
struct ShortcutProperties {
  enum IndividualProperties
  {
    PROPERTIES_TARGET                = 1U << 0,
    PROPERTIES_WORKING_DIR           = 1U << 1,
    PROPERTIES_ARGUMENTS             = 1U << 2,
    PROPERTIES_DESCRIPTION           = 1U << 3,
    PROPERTIES_ICON                  = 1U << 4,
    PROPERTIES_APP_ID                = 1U << 5,
    //PROPERTIES_DUAL_MODE             = 1U << 6,
    //PROPERTIES_TOAST_ACTIVATOR_CLSID = 1U << 7,
    // Be sure to update the values below when adding a new property.
    PROPERTIES_ALL = PROPERTIES_TARGET | PROPERTIES_WORKING_DIR | PROPERTIES_ARGUMENTS | PROPERTIES_DESCRIPTION | PROPERTIES_ICON | PROPERTIES_APP_ID/* |
                     PROPERTIES_DUAL_MODE | PROPERTIES_TOAST_ACTIVATOR_CLSID*/
  };

  void set_target(LPCTSTR target_in)
  {
    target = target_in;
    options |= PROPERTIES_TARGET;
  }

  void set_working_dir(LPCTSTR working_dir_in)
  {
    working_dir = working_dir_in;
    options |= PROPERTIES_WORKING_DIR;
  }

  void set_arguments(LPCTSTR arguments_in)
  {
    arguments = arguments_in;
    options |= PROPERTIES_ARGUMENTS;
  }

  void set_description(LPCTSTR description_in)
  {
    // Size restriction as per MSDN at http://goo.gl/OdNQq.
    // DCHECK_LE(description_in.size(), static_cast<size_t>(INFOTIPSIZE));
    description = description_in;
    options |= PROPERTIES_DESCRIPTION;
  }

  void set_icon(LPCTSTR icon_in, int icon_index_in)
  {
    icon       = icon_in;
    icon_index = icon_index_in;
    options |= PROPERTIES_ICON;
  }

  void set_app_id(LPCTSTR app_id_in)
  {
    app_id = app_id_in;
    options |= PROPERTIES_APP_ID;
  }

  //void set_dual_mode(bool dual_mode_in)
  //{
  //  dual_mode = dual_mode_in;
  //  options |= PROPERTIES_DUAL_MODE;
  //}

  //void set_toast_activator_clsid(const CLSID& toast_activator_clsid_in)
  //{
  //  toast_activator_clsid = toast_activator_clsid_in;
  //  options |= PROPERTIES_TOAST_ACTIVATOR_CLSID;
  //}

  // The target to launch from this shortcut. This is mandatory when creating
  // a shortcut.
  LPCTSTR target;
  // The name of the working directory when launching the shortcut.
  LPCTSTR working_dir;
  // The arguments to be applied to |target| when launching from this shortcut.
  LPCTSTR arguments;
  // The localized description of the shortcut.
  // The length of this string must be no larger than INFOTIPSIZE.
  LPCTSTR description;
  // The path to the icon (can be a dll or exe, in which case |icon_index| is
  // the resource id).
  LPCTSTR icon;
  int icon_index; // = -1;
  // The app model id for the shortcut.
  LPCTSTR app_id;
  // Whether this is a dual mode shortcut (Win8+).
  // bool dual_mode; // = false;
  // The CLSID of the COM object registered with the OS via the shortcut. This
  // is for app activation via user interaction with a toast notification in the
  // Action Center. (Win10 version 1607, build 14393, and beyond).
  // CLSID toast_activator_clsid;
  // Bitfield made of IndividualProperties. Properties set in |options| will be
  // set on the shortcut, others will be ignored.
  uint32_t options; // = 0U;
};

// This method creates (or updates) a shortcut link at |shortcut_path| using the
// information given through |properties|.
// Ensure you have initialized COM before calling into this function.
// |operation|: a choice from the ShortcutOperation enum.
// If |operation| is SHORTCUT_REPLACE_EXISTING or SHORTCUT_UPDATE_EXISTING and
// |shortcut_path| does not exist, this method is a no-op and returns false.
bool UpdateShortcutLink(LPCTSTR shortcut_path, const ShortcutProperties& properties);

} // namespace win
} // namespace base

#endif // BASE_WIN_SHORTCUT_H_
