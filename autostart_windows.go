package autostart

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

var startupDir string

func init() {
	startupDir = filepath.Join(os.Getenv("USERPROFILE"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
}

func (a *App) path() string {
	return filepath.Join(startupDir, a.Name+".lnk")
}

func (a *App) IsEnabled() bool {
	_, err := os.Stat(a.path())
	return err == nil
}

func (a *App) Enable() error {
	path := a.Exec[0]
	args := strings.Join(a.Exec[1:], " ")

	if err := os.MkdirAll(startupDir, 0777); err != nil {
		return err
	}
	err := CreateShortcut(a.path(), path, args, a.Icon)
	if err != nil {
		return fmt.Errorf("autostart: cannot create shortcut '%s' error: %s", a.path(), err.Error())
	}
	return nil
}

func (a *App) Disable() error {
	return os.Remove(a.path())
}

func CreateShortcut(path, target, args, iconPath string) error {

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	oleShellObject, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return err
	}
	defer oleShellObject.Release()
	wshell, err := oleShellObject.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer wshell.Release()
	// Shortcut path: path
	cs, err := oleutil.CallMethod(wshell, "CreateShortcut", path)
	if err != nil {
		return err
	}

	idispatch := cs.ToIDispatch()

	// Target: target
	_, err = oleutil.PutProperty(idispatch, "TargetPath", target)
	if err != nil {
		return err
	}

	// Arguments: args
	_, err = oleutil.PutProperty(idispatch, "Arguments", args)
	if err != nil {
		return err
	}

	// Icon path: iconPath
	if iconPath != "" {
		_, err = oleutil.PutProperty(idispatch, "IconLocation", iconPath)
		if err != nil {
			return err
		}
	}

	// save
	_, err = oleutil.CallMethod(idispatch, "Save")
	if err != nil {
		return err
	}
	return nil
}
