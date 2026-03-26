// Global using directives to resolve namespace conflicts between WPF and Windows Forms
// When UseWPF=true and UseWindowsForms=true, some types are ambiguous

global using Application = System.Windows.Application;
global using MessageBox = System.Windows.MessageBox;
global using Color = System.Windows.Media.Color;
global using Point = System.Windows.Point;
global using Size = System.Windows.Size;
global using Brush = System.Windows.Media.Brush;
global using Brushes = System.Windows.Media.Brushes;
global using Pen = System.Windows.Media.Pen;
global using FontFamily = System.Windows.Media.FontFamily;
global using Clipboard = System.Windows.Clipboard;
global using DataFormats = System.Windows.DataFormats;
global using DragDropEffects = System.Windows.DragDropEffects;
global using KeyEventArgs = System.Windows.Input.KeyEventArgs;
global using MouseEventArgs = System.Windows.Input.MouseEventArgs;
global using ContextMenu = System.Windows.Controls.ContextMenu;
global using MenuItem = System.Windows.Controls.MenuItem;
global using ToolTip = System.Windows.Controls.ToolTip;
global using Timer = System.Timers.Timer;
global using UserControl = System.Windows.Controls.UserControl;
global using HorizontalAlignment = System.Windows.HorizontalAlignment;
global using VerticalAlignment = System.Windows.VerticalAlignment;
global using Cursors = System.Windows.Input.Cursors;
global using Cursor = System.Windows.Input.Cursor;
global using IDataObject = System.Windows.IDataObject;
global using Orientation = System.Windows.Controls.Orientation;
global using ColorConverter = System.Windows.Media.ColorConverter;
