using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using NewUserAutomation.App.ViewModels;

namespace NewUserAutomation.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel();
        DataContext = _viewModel;
    }

    private void OnBrowseFileClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Form files (*.txt;*.docx)|*.txt;*.docx|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            _viewModel.SetSelectedFilePath(dialog.FileName);
        }
    }

    private void OnInnerListBoxPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not ListBox listBox)
        {
            return;
        }

        var innerScroller = FindChildScrollViewer(listBox);
        if (innerScroller is not null)
        {
            var atTop = innerScroller.VerticalOffset <= 0;
            var atBottom = innerScroller.VerticalOffset >= innerScroller.ScrollableHeight;

            if ((e.Delta > 0 && !atTop) || (e.Delta < 0 && !atBottom))
            {
                ScrollWithWheel(innerScroller, e.Delta);
                e.Handled = true;
                return;
            }
        }

        var parentScroller = FindParentScrollViewer((DependencyObject)sender);
        if (parentScroller is null)
        {
            return;
        }

        ScrollWithWheel(parentScroller, e.Delta);
        e.Handled = true;
    }

    private void OnLiveRunListPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not ListBox listBox)
        {
            return;
        }

        var innerScroller = FindChildScrollViewer(listBox);
        if (innerScroller is not null)
        {
            var atTop = innerScroller.VerticalOffset <= 0;
            var atBottom = innerScroller.VerticalOffset >= innerScroller.ScrollableHeight;

            if ((e.Delta > 0 && !atTop) || (e.Delta < 0 && !atBottom))
            {
                innerScroller.ScrollToVerticalOffset(innerScroller.VerticalOffset - e.Delta);
                e.Handled = true;
                return;
            }
        }

        OnInnerListBoxPreviewMouseWheel(sender, e);
    }

    private static ScrollViewer? FindParentScrollViewer(DependencyObject? child)
    {
        var current = child;
        while (current is not null)
        {
            if (current is ScrollViewer scrollViewer)
            {
                return scrollViewer;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static ScrollViewer? FindChildScrollViewer(DependencyObject? parent)
    {
        if (parent is null)
        {
            return null;
        }

        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is ScrollViewer scrollViewer)
            {
                return scrollViewer;
            }

            var nested = FindChildScrollViewer(child);
            if (nested is not null)
            {
                return nested;
            }
        }

        return null;
    }

    private static void ScrollWithWheel(ScrollViewer scroller, int mouseWheelDelta)
    {
        var wheelNotches = mouseWheelDelta / 120.0;
        var lines = SystemParameters.WheelScrollLines;
        if (lines <= 0)
        {
            lines = 3;
        }

        var pixelsPerLine = 18.0;
        var delta = wheelNotches * lines * pixelsPerLine;
        var targetOffset = scroller.VerticalOffset - delta;
        if (targetOffset < 0)
        {
            targetOffset = 0;
        }
        if (targetOffset > scroller.ScrollableHeight)
        {
            targetOffset = scroller.ScrollableHeight;
        }

        scroller.ScrollToVerticalOffset(targetOffset);
    }
}