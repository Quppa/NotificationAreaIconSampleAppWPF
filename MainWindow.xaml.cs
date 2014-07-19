namespace NotifyIconSample
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Interop;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Navigation;
    using System.Windows.Shapes;
    using Microsoft.Win32;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            this.Pinned = false;
            this.MouseClickToHideNotifyIcon = false;

            this.PreferenceEventHandler = null;
            this.PreferenceEvent = null;

            this.CreateNotifyIcon();
        }

        /// <summary>
        /// Delegate for handling user preference changes (namely desktop preference changes).
        /// </summary>
        /// <param name="sender">The source of the event. When this event is raised by the SystemEvents class, this object is always null.</param>
        /// <param name="e">A UserPreferenceChangedEventArgs that contains the event data.</param>
        private delegate void UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e);

        /// <summary>
        /// User preferences changed event handler. Set when the window is made visible and unset when the window is hidden.
        /// </summary>
        private event UserPreferenceChanged PreferenceEvent;

        /// <summary>
        /// Gets or sets a handler for user preference changes.
        /// </summary>
        private UserPreferenceChangedEventHandler PreferenceEventHandler { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not the window is pinned open.
        /// </summary>
        private bool Pinned { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the window's notify icon.
        /// </summary>
        private System.Windows.Forms.NotifyIcon NotifyIcon { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user hid the window by clicking the notify icon a second time.
        /// </summary>
        private bool MouseClickToHideNotifyIcon { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the location of the cursor when the window was last hidden by clicking the notify icon a second time.
        /// </summary>
        private Point MouseClickToHideNotifyIconPoint { get; set; }

        /// <summary>
        /// Updates the display (position and appearance) of the window if it is currently visible.
        /// </summary>
        /// <param name="activatewindow">True if the window should be activated, false if not.</param>
        public void UpdateWindowDisplayIfOpen(bool activatewindow)
        {
            if (this.Visibility == Visibility.Visible)
                this.UpdateWindowDisplay(activatewindow);
        }

        /// <summary>
        /// Updates the display (position and appearance) of the window.
        /// </summary>
        /// <param name="activatewindow">True if the window should be activated, false if not.</param>
        public void UpdateWindowDisplay(bool activatewindow)
        {
            if (this.IsLoaded)
            {
                // set handlers if necessary
                this.SetHandlers();

                // get the handle of the window
                HwndSource windowhandlesource = PresentationSource.FromVisual(this) as HwndSource;

                bool glassenabled = Compatibility.IsDWMEnabled;

                //// update location

                Rect windowbounds = (glassenabled ? WindowPositioning.GetWindowSize(windowhandlesource.Handle) : WindowPositioning.GetWindowClientAreaSize(windowhandlesource.Handle));

                // work out the current screen's DPI
                Matrix screenmatrix = windowhandlesource.CompositionTarget.TransformToDevice;

                double dpiX = screenmatrix.M11; // 1.0 = 96 dpi
                double dpiY = screenmatrix.M22; // 1.25 = 120 dpi, etc.

                Point position = WindowPositioning.GetWindowPosition(this.NotifyIcon, windowbounds.Width, windowbounds.Height, dpiX, this.Pinned);

                // translate wpf points to screen coordinates
                Point screenposition = new Point(position.X / dpiX, position.Y / dpiY);

                this.Left = screenposition.X;
                this.Top = screenposition.Y;

                // update borders
                if (glassenabled)
                    this.Style = (Style)FindResource("AeroBorderStyle");
                else
                    this.SetNonGlassBorder(this.IsActive);

                // fix aero border if necessary
                if (glassenabled)
                {
                    // set the root border element's margin to 1 pixel
                    WindowBorder.Margin = new Thickness(1 / dpiX);
                    this.BorderThickness = new Thickness(0);

                    // set the background of the window to transparent (otherwise the inner border colour won't be visible)
                    windowhandlesource.CompositionTarget.BackgroundColor = Colors.Transparent;

                    // get dpi-dependent aero border width
                    int xmargin = Convert.ToInt32(1); // 1 unit wide
                    int ymargin = Convert.ToInt32(1); // 1 unit tall

                    NativeMethods.MARGINS margins = new NativeMethods.MARGINS() { cxLeftWidth = xmargin, cxRightWidth = xmargin, cyBottomHeight = ymargin, cyTopHeight = ymargin };

                    NativeMethods.DwmExtendFrameIntoClientArea(windowhandlesource.Handle, ref margins);
                }
                else
                {
                    WindowBorder.Margin = new Thickness(0); // reset the margin if the DWM is disabled
                    this.BorderThickness = new Thickness(1 / dpiX); // set the window's border thickness to 1 pixel
                }

                if (activatewindow)
                {
                    this.Show();
                    this.Activate();
                }
            }
        }

        #region Event Handlers

        /// <summary>
        /// Adds hook to custom DefWindowProc function.
        /// </summary>
        /// <param name="e">OnSourceInitialized EventArgs.</param>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(this.WndProc);
        }

        /// <summary>
        /// Custom DefWindowProc function used to disable resize and update the window appearance when the window size is changed or the DWM is enabled/disabled.
        /// </summary>
        /// <param name="hWnd">A handle to the window procedure that received the message.</param>
        /// <param name="msg">The message.</param>
        /// <param name="wParam">Additional message information. The content of this parameter depends on the value of the Msg parameter (wParam).</param>
        /// <param name="lParam">Additional message information. The content of this parameter depends on the value of the Msg parameter (lParam).</param>
        /// <param name="handled">True if the message has been handled by the custom window procedure, false if not.</param>
        /// <returns>The return value is the result of the message processing and depends on the message.</returns>
        private IntPtr WndProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (this.IsLoaded && this.Visibility == Visibility.Visible)
            {
                switch (msg)
                {
                    case NativeMethods.WM_NCHITTEST:
                        // if the mouse pointer is not over the client area of the tab
                        // ignore it - this disables resize on the glass chrome
                        if (!NativeMethods.IsOverClientArea(hWnd, wParam, lParam))
                            handled = true;

                        break;

                    case NativeMethods.WM_SETCURSOR:
                        if (!NativeMethods.IsOverClientArea(hWnd, wParam, lParam))
                        {
                            // the high word of lParam specifies the mouse message identifier
                            // we only want to handle mouse down messages on the border
                            int hiword = (int)lParam >> 16;
                            if (hiword == NativeMethods.WM_LBUTTONDOWN
                                || hiword == NativeMethods.WM_RBUTTONDOWN
                                || hiword == NativeMethods.WM_MBUTTONDOWN
                                || hiword == NativeMethods.WM_XBUTTONDOWN)
                            {
                                handled = true;
                                this.Focus(); // focus the window
                            }
                        }

                        break;

                    case NativeMethods.WM_DWMCOMPOSITIONCHANGED:
                        // update window appearance accordingly
                        this.UpdateWindowDisplayIfOpen(false);

                        break;

                    case NativeMethods.WM_SIZE:
                        // update window appearance accordingly
                        this.UpdateWindowDisplayIfOpen(false);

                        break;
                }
            }

            return IntPtr.Zero;
        }

        /// <summary>
        /// Sets handlers for notifying the application of desktop preference changes (taskbar movements, etc.).
        /// </summary>
        private void SetHandlers()
        {
            if (this.PreferenceEvent == null && this.PreferenceEventHandler == null)
            {
                this.PreferenceEvent = new UserPreferenceChanged(this.DesktopPreferenceChangedHandler);
                this.PreferenceEventHandler = new UserPreferenceChangedEventHandler(this.PreferenceEvent);
                SystemEvents.UserPreferenceChanged += this.PreferenceEventHandler;
            }
        }

        /// <summary>
        /// Releases handlers set by SetHandlers().
        /// </summary>
        private void ReleaseHandlers()
        {
            if (this.PreferenceEvent != null || this.PreferenceEventHandler != null)
            {
                SystemEvents.UserPreferenceChanged -= this.PreferenceEventHandler;
                this.PreferenceEvent = null;
                this.PreferenceEventHandler = null;
            }
        }

        /// <summary>
        /// Handler for UserPreferenceChangedEventArgs that updates the window display when the user modifies his or her desktop.
        /// Note: This does not detect taskbar changes when the taskbar is set to auto-hide.
        /// </summary>
        /// <param name="sender">The source of the event. When this event is raised by the SystemEvents class, this object is always null.</param>
        /// <param name="e">A UserPreferenceChangedEventArgs that contains the event data.</param>
        private void DesktopPreferenceChangedHandler(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Desktop)
                this.UpdateWindowDisplayIfOpen(false);
        }

        #endregion

        /// <summary>
        /// Notify icon clicked or double-clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="args">System.Windows.Forms.MouseEventArgs (which mouse button was pressed, etc.).</param>
        private void NotifyIconClick(object sender, System.Windows.Forms.MouseEventArgs args)
        {
            if (!this.IsLoaded)
                this.Show();
                //this.Visibility = Visibility.Visible;
            
            if (args.Button == System.Windows.Forms.MouseButtons.Left &&
                (!this.MouseClickToHideNotifyIcon
                || (WindowPositioning.GetCursorPosition().X != this.MouseClickToHideNotifyIconPoint.X || WindowPositioning.GetCursorPosition().Y != this.MouseClickToHideNotifyIconPoint.Y)))
                this.UpdateWindowDisplay(true);
            else
                this.MouseClickToHideNotifyIcon = false;
        }

        /// <summary>
        /// Sets the border of the window when the DWM is not enabled. The colour of the border depends on whether the window is active or not.
        /// </summary>
        /// <param name="windowactivated">True if the window is active, false if not.</param>
        private void SetNonGlassBorder(bool windowactivated)
        {
            if (windowactivated)
                this.Style = (Style)FindResource("ClassicBorderStyle");
            else
                this.Style = (Style)FindResource("ClassicBorderStyleInactive");
        }

        /// <summary>
        /// Window deactivated method. Hides window if not pinned and sets deactivated border colour if DWM is disabled.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Event arguments.</param>
        private void Window_Deactivated(object sender, EventArgs e)
        {
            this.HideWindowIfNotPinned();

            if (!Compatibility.IsDWMEnabled)
                this.SetNonGlassBorder(false);
        }
        
        /// <summary>
        /// Hides the window if it is not pinned open.
        /// </summary>
        private void HideWindowIfNotPinned()
        {
            if (!this.Pinned)
                this.HideWindow();
        }

        /// <summary>
        /// Hides the window
        /// </summary>
        private void HideWindow()
        {
            // note if mouse is over the notify icon when hiding the window
            // if it is, we will assume that the user clicked the icon to hide the window
            this.MouseClickToHideNotifyIcon = WindowPositioning.IsCursorOverNotifyIcon(this.NotifyIcon) && WindowPositioning.IsNotificationAreaActive;
            if (this.MouseClickToHideNotifyIcon)
                this.MouseClickToHideNotifyIconPoint = WindowPositioning.GetCursorPosition();

            this.ReleaseHandlers();
            this.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// Window closing method. Disposes of the notify icon, removes the custom window procedure and releases user preference change handlers.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Cancel event arguments.</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // remove the notify icon
            this.NotifyIcon.Visible = false;
            this.NotifyIcon.Dispose();

            if (this.IsLoaded)
            {
                HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
                source.RemoveHook(this.WndProc);
            }

            this.ReleaseHandlers();
        }

        /// <summary>
        /// Exit button clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Routed event arguments.</param>
        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Exit();
        }

        /// <summary>
        /// Exit menu button (notify icon) clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Event arguments.</param>
        private void ExitMenuEventHandler(object sender, EventArgs e)
        {
            this.Exit();
        }

        /// <summary>
        /// Close the application.
        /// </summary>
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        /// <summary>
        /// Pin button clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Routed event arguments.</param>
        private void PinButton_Click(object sender, RoutedEventArgs e)
        {
            this.SetPin(true);
        }

        /// <summary>
        /// Pin/Unpin menu button (notify icon) clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Event arguments.</param>
        private void PinMenuEventHandler(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem menuitem = sender as System.Windows.Forms.MenuItem;
            if (menuitem != null)
                menuitem.Checked = !menuitem.Checked;
            this.ReversePin();
        }

        /// <summary>
        /// Reverses the current pin state.
        /// </summary>
        private void ReversePin()
        {
            this.SetPin(!this.Pinned);
        }

        /// <summary>
        /// Sets the window's pin status and updates the window display or hides the window.
        /// </summary>
        /// <param name="pinned">True if the window is to be pinned open, false if it is to be unpinned.</param>
        private void SetPin(bool pinned)
        {
            this.Pinned = pinned;

            if (pinned)
                this.UpdateWindowDisplay(true);
            else
                this.HideWindow();
        }

        /// <summary>
        /// Window activated method. Sets the window border colour if the DWM is disabled.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Event arguments.</param>
        private void Window_Activated(object sender, EventArgs e)
        {
            if (!Compatibility.IsDWMEnabled)
                this.SetNonGlassBorder(true);
        }

        /// <summary>
        /// Window loaded method. We update the window display before it is made visible, otherwise the user will see its position jump when the notify icon is first clicked.
        /// </summary>
        /// <param name="sender">Sender of the message.</param>
        /// <param name="e">Routed event arguments.</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.UpdateWindowDisplay(false);
        }

        /// <summary>
        /// Creates notify icon and menu.
        /// </summary>
        private void CreateNotifyIcon()
        {
            System.Windows.Forms.NotifyIcon notifyicon;

            notifyicon = new System.Windows.Forms.NotifyIcon();

            // set icon
            Stream iconstream = Application.GetResourceStream(new Uri("pack://application:,,,/NotifyIconSample;component/Tray.ico")).Stream;
            notifyicon.Icon = new System.Drawing.Icon(iconstream, System.Windows.Forms.SystemInformation.SmallIconSize);
            iconstream.Close();

            System.Windows.Forms.MenuItem mnuPin = new System.Windows.Forms.MenuItem("Pin", new EventHandler(this.PinMenuEventHandler));
            System.Windows.Forms.MenuItem mnuExit = new System.Windows.Forms.MenuItem("Exit", new EventHandler(this.ExitMenuEventHandler));
            
            System.Windows.Forms.MenuItem[] menuitems = new System.Windows.Forms.MenuItem[]
            {
                mnuPin, new System.Windows.Forms.MenuItem("-"), mnuExit
            };

            System.Windows.Forms.ContextMenu contextmenu = new System.Windows.Forms.ContextMenu(menuitems);

            notifyicon.ContextMenu = contextmenu;

            notifyicon.MouseClick += this.NotifyIconClick;
            notifyicon.MouseDoubleClick += this.NotifyIconClick;

            notifyicon.Visible = true;

            this.NotifyIcon = notifyicon;
        }

        /// <summary>
        /// Hyperlink clicked method.
        /// </summary>
        /// <param name="sender">The sender of the message.</param>
        /// <param name="e">Routed event arguments.</param>
        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.quppa.net");
        }
    }
}
