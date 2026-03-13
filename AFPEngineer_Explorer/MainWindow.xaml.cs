using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using AFPEngineer_Explorer.Services;
using OpenAI.Chat;
using System.ClientModel;

namespace AFPEngineer_Explorer
{
    public class AfpNode : System.ComponentModel.INotifyPropertyChanged
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string FriendlyName { get; set; }
        public string Url { get; set; }
        public long Offset { get; set; }
        public int Length { get; set; }
        public int? RepeatingGroupLength { get; set; }
        public string DisplayName => $"{Id} - {FriendlyName} ({Name})";
        public ObservableCollection<AfpNode> Children { get; set; } = new ObservableCollection<AfpNode>();
        public AfpNode Parent { get; set; }

        private bool _isExpanded;
        public bool IsExpanded 
        { 
            get => _isExpanded; 
            set { _isExpanded = value; OnPropertyChanged(nameof(IsExpanded)); } 
        }

        private bool _isSelected;
        public bool IsSelected 
        { 
            get => _isSelected; 
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } 
        }

        private System.Windows.Media.Brush _foregroundBrush = System.Windows.Media.Brushes.Black;
        public System.Windows.Media.Brush ForegroundBrush
        {
            get => _foregroundBrush;
            set { _foregroundBrush = value; OnPropertyChanged(nameof(ForegroundBrush)); }
        }

        private System.Windows.Media.Brush _backgroundBrush = System.Windows.Media.Brushes.Transparent;
        public System.Windows.Media.Brush BackgroundBrush
        {
            get => _backgroundBrush;
            set { _backgroundBrush = value; OnPropertyChanged(nameof(BackgroundBrush)); }
        }

        private FontWeight _fontWeight = FontWeights.Normal;
        public FontWeight FontWeight
        {
            get => _fontWeight;
            set { _fontWeight = value; OnPropertyChanged(nameof(FontWeight)); }
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string prop) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(prop));
    }

    public partial class MainWindow : Window
    {
        // EBCDIC (Code Page 37) is the standard for most AFP files
        private static Encoding ebcdic;
        
        public MainWindow()
        {
            // Add this line to unlock EBCDIC support
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            ebcdic = Encoding.GetEncoding("IBM037");
            try
            {
                string definitionsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "definitions", "structured_fields.json");
                MappingService.Instance.LoadDefinitions(definitionsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load definitions: {ex.Message}");
            }

            InitializeComponent();
            // Bypass the generator: Find the tree manually and hook up the event
            var tree = this.FindName("AfpTreeView") as TreeView;
            if (tree != null)
            {
                tree.SelectedItemChanged += AfpTreeView_SelectedItemChanged;
            }
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "AFP Files (*.afp;*.out;*.prn)|*.afp;*.out;*.prn|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                ParseAfpFile(openFileDialog.FileName);
            }
        }

        private void SearchPrevBtn_Click(object sender, RoutedEventArgs e)
        {
            SearchAction(-1);
        }

        private void SearchNextBtn_Click(object sender, RoutedEventArgs e)
        {
            int direction = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift ? -1 : 1;
            SearchAction(direction);
        }

        private int _lastSearchIndex = -1;
        private string _lastSearchTerm = "";
        private List<AfpNode> _flatNodeList = new List<AfpNode>();

        private void SearchAction(int direction)
        {
            var searchBox = this.FindName("SearchBox") as TextBox;
            if (searchBox == null || string.IsNullOrWhiteSpace(searchBox.Text)) return;

            string term = searchBox.Text.ToLower();

            if (term != _lastSearchTerm)
            {
                // Clear previous highlight before rebuilding list
                if (_flatNodeList != null && _lastSearchIndex >= 0 && _lastSearchIndex < _flatNodeList.Count)
                {
                    _flatNodeList[_lastSearchIndex].ForegroundBrush = System.Windows.Media.Brushes.Black;
                    _flatNodeList[_lastSearchIndex].BackgroundBrush = System.Windows.Media.Brushes.Transparent;
                    _flatNodeList[_lastSearchIndex].FontWeight = FontWeights.Normal;
                }

                _lastSearchTerm = term;
                _lastSearchIndex = -1;
                // Rebuild flat list
                _flatNodeList.Clear();
                var tree = this.FindName("AfpTreeView") as TreeView;
                if (tree?.ItemsSource is IEnumerable<AfpNode> roots)
                {
                    foreach (var root in roots) FlattenNodes(root, _flatNodeList);
                }
            }

            if (_flatNodeList.Count == 0) return;

            // Clear previous highlight (in case we didn't change term)
            if (_lastSearchIndex >= 0 && _lastSearchIndex < _flatNodeList.Count)
            {
                _flatNodeList[_lastSearchIndex].ForegroundBrush = System.Windows.Media.Brushes.Black;
                _flatNodeList[_lastSearchIndex].BackgroundBrush = System.Windows.Media.Brushes.Transparent;
                _flatNodeList[_lastSearchIndex].FontWeight = FontWeights.Normal;
            }

            // Move pointer
            if (_lastSearchIndex == -1) _lastSearchIndex = direction > 0 ? 0 : _flatNodeList.Count - 1;
            else _lastSearchIndex += direction;

            bool found = false;

            for (int i = 0; i < _flatNodeList.Count; i++)
            {
                // Navigate wrapped around the list based on direction
                int checkIndex = ((_lastSearchIndex + (i * direction)) % _flatNodeList.Count + _flatNodeList.Count) % _flatNodeList.Count;
                var node = _flatNodeList[checkIndex];
                if ((node.Name != null && node.Name.ToLower().Contains(term)) ||
                    (node.FriendlyName != null && node.FriendlyName.ToLower().Contains(term)) ||
                    (node.Id != null && node.Id.ToLower().Contains(term)))
                {
                    _lastSearchIndex = checkIndex;
                    found = true;
                    node.ForegroundBrush = System.Windows.Media.Brushes.White;
                    node.BackgroundBrush = System.Windows.Media.Brushes.DarkRed;
                    node.FontWeight = FontWeights.Bold;
                    node.IsSelected = true;

                    // Expand parents
                    var parent = node.Parent;
                    while (parent != null)
                    {
                        parent.IsExpanded = true;
                        parent = parent.Parent;
                    }

                    AfpTreeView.UpdateLayout();
                    BringNodeIntoView(node, AfpTreeView);
                    break;
                }
            }

            if (!found)
            {
                MessageBox.Show("No matching structured field found.");
                _lastSearchIndex = -1;
            }
        }

        private void FlattenNodes(AfpNode node, List<AfpNode> list)
        {
            list.Add(node);
            foreach (var child in node.Children)
                FlattenNodes(child, list);
        }

        private void BringNodeIntoView(AfpNode targetNode, TreeView tree)
        {
            if (targetNode == null || tree == null) return;
            var path = new List<AfpNode>();
            var p = targetNode;
            while(p != null) { path.Insert(0, p); p = p.Parent; }
            
            ItemsControl currentContainer = tree;
            foreach(var node in path)
            {
                if (currentContainer == null) break;

                // Make sure container is expanded
                if (currentContainer is TreeViewItem tvi && !tvi.IsExpanded)
                {
                    tvi.IsExpanded = true;
                    currentContainer.UpdateLayout();
                }

                // Force generation of children
                var generator = currentContainer.ItemContainerGenerator;
                if (generator.Status != System.Windows.Controls.Primitives.GeneratorStatus.ContainersGenerated)
                {
                    currentContainer.ApplyTemplate();
                    var itemsPresenter = (ItemsPresenter)currentContainer.Template.FindName("ItemsHost", currentContainer);
                    if (itemsPresenter != null) itemsPresenter.ApplyTemplate();
                    currentContainer.UpdateLayout();
                }

                var container = generator.ContainerFromItem(node) as TreeViewItem;
                if (container == null)
                {
                    currentContainer.UpdateLayout();
                    container = generator.ContainerFromItem(node) as TreeViewItem;
                }
                
                if (node == targetNode && container != null)
                {
                    container.BringIntoView();
                    // Don't focus, so search box keeps focus
                }
                currentContainer = container;
            }
        }

        private string _currentFilePath;

        private async void ParseAfpFile(string filePath)
        {
            var tree = this.FindName("AfpTreeView") as TreeView;
            if (tree == null) return;

            tree.ItemsSource = null;
            _currentFilePath = filePath;
            
            try
            {
                // Offload the basic sequential structure scanning to a background thread
                var nodes = await Task.Run(() => BuildAfpHierarchy(filePath));
                tree.ItemsSource = nodes;

                _flatNodeList.Clear();
                foreach (var root in nodes) FlattenNodes(root, _flatNodeList);

                // Document Display Mapping 
                _pageNodes.Clear();
                foreach (var node in _flatNodeList)
                {
                    if (node.Name == "BPG")
                        _pageNodes.Add(node);
                }
                
                _totalDocPages = Math.Max(1, _pageNodes.Count);
                _currentDocPage = 1;
                UpdateDocumentDisplay();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private ObservableCollection<AfpNode> BuildAfpHierarchy(string filePath)
        {
            var rootNodes = new ObservableCollection<AfpNode>();
            Stack<AfpNode> containerStack = new Stack<AfpNode>();
            int? currentFniRgLen = null;
            int? currentFnoRgLen = null;
            int? currentFnpRgLen = null;
            int? currentFnmRgLen = null;
            int? currentCpiRgLen = null;

            using (BinaryReader reader = new BinaryReader(File.OpenRead(filePath)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    if (reader.ReadByte() == 0x5A)
                    {
                        byte hi = reader.ReadByte();
                        byte lo = reader.ReadByte();
                        ushort length = (ushort)((hi << 8) | lo);
                        byte[] idBytes = reader.ReadBytes(3);
                        string hexId = BitConverter.ToString(idBytes).Replace("-", "");
                        
                        // Payload length excludes the 8-byte introducer portion (Length(2) + ID(3) + Flag(1) + Seq(2))
                        int payloadLength = length - 8;
                        long fileOffset = reader.BaseStream.Position - 6;

                        // Check for FNC to capture repeating group lengths
                        if (hexId == "D3A789" && payloadLength >= 22)
                        {
                            long savePos = reader.BaseStream.Position;
                            reader.BaseStream.Seek(3 + 14, SeekOrigin.Current); // skip 3 bytes header remainder + 14 bytes to FNORGLen
                            currentFnoRgLen = reader.ReadByte();
                            currentFniRgLen = reader.ReadByte();
                            reader.BaseStream.Seek(4, SeekOrigin.Current); // skip to FNPRGLen
                            currentFnpRgLen = reader.ReadByte();
                            currentFnmRgLen = reader.ReadByte();
                            reader.BaseStream.Seek(savePos, SeekOrigin.Begin);
                        }

                        // Check for CPC to capture CPIRGLen
                        if (hexId == "D3A787" && payloadLength >= 10)
                        {
                            long savePos = reader.BaseStream.Position;
                            reader.BaseStream.Seek(3 + 9, SeekOrigin.Current); // skip 3 bytes header remainder + 9 bytes to CPIRGLen
                            currentCpiRgLen = reader.ReadByte();
                            reader.BaseStream.Seek(savePos, SeekOrigin.Begin);
                        }

                        // Check for TLE to extract Triplet 0x02 and 0x36
                        string tleAddendum = "";
                        if (hexId == "D3A090" && payloadLength > 0)
                        {
                            long savePos = reader.BaseStream.Position;
                            reader.BaseStream.Seek(3, SeekOrigin.Current); // skip 3 bytes header remainder
                            byte[] tlePayload = reader.ReadBytes(payloadLength);
                            string attrName = "";
                            string attrValue = "";
                            
                            int idx = 0;
                            while (idx < tlePayload.Length)
                            {
                                int tLen = tlePayload[idx];
                                if (tLen < 2 || idx + tLen > tlePayload.Length) break;
                                byte tId = tlePayload[idx + 1];
                                
                                if (tId == 0x02 && tLen >= 4) // Fully Qualified Name
                                {
                                    attrName = ebcdic.GetString(tlePayload, idx + 4, tLen - 4).Trim();
                                }
                                else if (tId == 0x36 && tLen >= 4) // Attribute Value
                                {
                                    attrValue = ebcdic.GetString(tlePayload, idx + 3, tLen - 3).Trim();
                                }
                                idx += tLen;
                            }
                            
                            if (!string.IsNullOrEmpty(attrName) || !string.IsNullOrEmpty(attrValue))
                            {
                                tleAddendum = $" - {attrName}: {attrValue}";
                            }
                            reader.BaseStream.Seek(savePos, SeekOrigin.Begin);
                        }

                        var sfInfo = LookupSfInfo(hexId);

                        int? rgLen = null;
                        if (hexId == "D38C89") rgLen = currentFniRgLen;
                        else if (hexId == "D38C8D") rgLen = currentFnoRgLen; // FNO
                        else if (hexId == "D3B189") rgLen = currentFnpRgLen; // FNP
                        else if (hexId == "D3A289") rgLen = currentFnmRgLen; // FNM
                        else if (hexId == "D38C87") rgLen = currentCpiRgLen; // CPI

                        AfpNode newNode = new AfpNode
                        {
                            Id = hexId,
                            Name = sfInfo.Acronym,
                            FriendlyName = string.IsNullOrEmpty(tleAddendum) ? sfInfo.FriendlyName : sfInfo.FriendlyName + tleAddendum,
                            Url = GetDefinitionUrl(sfInfo.Acronym),
                            Offset = fileOffset,
                            Length = payloadLength,
                            RepeatingGroupLength = rgLen
                        };

                        // Fast Forward to the next record (skip the remaining introducer bytes + payload)
                        reader.BaseStream.Seek(3 + payloadLength, SeekOrigin.Current);

                        // Nesting Logic
                        if (hexId.StartsWith("D3A8")) // Begin elements
                        {
                            if (containerStack.Count > 0) 
                            {
                                newNode.Parent = containerStack.Peek();
                                containerStack.Peek().Children.Add(newNode);
                            }
                            else rootNodes.Add(newNode);
                            containerStack.Push(newNode);
                        }
                        else if (hexId.StartsWith("D3A9")) // End elements
                        {
                            if (containerStack.Count > 0) containerStack.Pop();
                            if (containerStack.Count > 0) 
                            {
                                newNode.Parent = containerStack.Peek();
                                containerStack.Peek().Children.Add(newNode);
                            }
                            else rootNodes.Add(newNode);
                        }
                        else
                        {
                            if (containerStack.Count > 0) 
                            {
                                newNode.Parent = containerStack.Peek();
                                containerStack.Peek().Children.Add(newNode);
                            }
                            else rootNodes.Add(newNode);
                        }
                    }
                }
            }
            return rootNodes;
        }

        private SfInfo LookupSfInfo(string id)
        {
            var sfInfo = AfpConstants.GetSfInfo(id);
            if (sfInfo.Acronym == "UNKNOWN")
            {
                // Fallback attempt to see if it's found in mapping service dynamically, else generic
                return new SfInfo { Acronym = "UNK", FriendlyName = "Unknown Structured Field" };
            }
            return sfInfo;
        }

        private string GetDefinitionUrl(string acronym)
        {
            if (acronym == "UNK" || acronym == "UNKNOWN") return null;
            
            var def = MappingService.Instance.GetDefinition(acronym);
            if (def != null && def.Definition != null)
            {
                if (def.Definition.TryGetValue("AFPCCManual", out var urlObj))
                {
                    return urlObj.ToString();
                }
            }
            return null;
        }

        private string GetExpertAdvice(string id) => id switch
        {
            "D3A6AF" => "PGD (Page Descriptor): Defines the page size. The first 3 bytes after the ID define the 'Units per 10 inches'. Usually, it's 14400 (1440 dpi).",
            "D3A6BB" => "PTD (Presentation Text Descriptor): This sets the default font and color for all the text on this page.",
            "D3EEBB" => "PTX (Presentation Text): This is where the magic happens! It contains PTOCA control sequences like 'AMB' (Absolute Move Baseline) to position text.",
            "D3ACAF" => "PAG (Page Attribute): Contains metadata about the page, like which side of the paper to print on (Duplex).",
            "D3A8C7" => "BMO (Begin Overlay): This is an 'electronic form' that gets placed underneath your data.",
            "D3A8AD" => "BDI (Begin Document Index): This is used by software like IBM Content Manager to find specific pages (like an Account Number) quickly.",
            _ => "This is a MO:DCA Structured Field. Refer to the 'AFP Consortium Mixed Object Document Content Architecture' PDF for the bit-level breakdown of the triplets."
        };
        private void ListView_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (!e.Handled)
            {
                e.Handled = true;
                var eventArg = new System.Windows.Input.MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
                eventArg.RoutedEvent = UIElement.MouseWheelEvent;
                eventArg.Source = sender;
                var parent = ((Control)sender).Parent as UIElement;
                if (parent != null)
                {
                    parent.RaiseEvent(eventArg);
                }
            }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            if (e.Uri != null)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                });
                e.Handled = true;
            }
        }

        private GridViewColumn CreateGridColumn(string header, string bindPath, bool isMultiLine = false)
        {
            var col = new GridViewColumn { Header = header };
            
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(bindPath));
            factory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            if(isMultiLine) factory.SetValue(TextBlock.MarginProperty, new Thickness(5));

            var tmpl = new DataTemplate { VisualTree = factory };
            col.CellTemplate = tmpl;

            return col;
        }

        private string FormatHexDump(byte[] payload, long offset)
        {
            var sb = new StringBuilder();
            int bytesPerLine = 16;
            
            for (int i = 0; i < payload.Length; i += bytesPerLine)
            {
                sb.Append($"0x{(offset + i):X8}: ");
                
                // Hex values
                for (int j = 0; j < bytesPerLine; j++)
                {
                    if (i + j < payload.Length)
                        sb.Append($"{payload[i + j]:X2} ");
                    else
                        sb.Append("   ");
                }
                
                sb.Append("  ");
                
                // ASCII/EBCDIC representation
                byte[] lineBytes = new byte[Math.Min(bytesPerLine, payload.Length - i)];
                Array.Copy(payload, i, lineBytes, 0, lineBytes.Length);
                string text = ebcdic.GetString(lineBytes);
                text = Regex.Replace(text, @"[^\u0020-\u007E\u00A0-\u00FF]", ".");
                sb.Append(text);
                
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private async void AfpTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var fieldNameTxt = this.FindName("FieldNameTxt") as TextBlock;
            var exploreTitleTxt = this.FindName("ExploreTitleTxt") as TextBlock;
            var exploreDocLinkBlock = this.FindName("ExploreDocLinkBlock") as TextBlock;
            var exploreLink = this.FindName("ExploreLink") as Hyperlink;
            var exploreOffsetTxt = this.FindName("ExploreOffsetTxt") as TextBlock;
            var exploreLengthTxt = this.FindName("ExploreLengthTxt") as TextBlock;
            var exploreListView = this.FindName("ExploreListView") as ListView;
            var exploreFallbackTxt = this.FindName("ExploreFallbackTxt") as TextBlock;
            var hexdumpHeaderTxt = this.FindName("HexdumpHeaderTxt") as TextBlock;
            var hexdumpTxt = this.FindName("HexdumpTxt") as TextBox;
            var viewImageBtn = this.FindName("ViewImageBtn") as Button;
            var askAiBtn = this.FindName("AskAIBtn") as Button;
            var aiResponseTxt = this.FindName("AiResponseTxt") as TextBox;
            var tree = sender as TreeView;
            var exploreTab = this.FindName("ExploreTab") as TabItem;
            var imageViewerTab = this.FindName("ImageViewerTab") as TabItem;
            var loadingOverlay = this.FindName("LoadingOverlay") as Grid;

            if (exploreTab != null) exploreTab.IsSelected = true;
            if (imageViewerTab != null) imageViewerTab.Visibility = Visibility.Collapsed;

            if (tree?.SelectedItem is AfpNode selected && !string.IsNullOrEmpty(_currentFilePath))
            {
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Visible;

                if (askAiBtn != null)
                {
                    askAiBtn.Visibility = Visibility.Visible;
                    askAiBtn.Tag = selected;
                }
                if (aiResponseTxt != null) 
                    aiResponseTxt.Visibility = Visibility.Collapsed;

                if (fieldNameTxt != null) fieldNameTxt.Visibility = Visibility.Collapsed; // Replaced by Explore Title
                if (exploreTitleTxt != null) exploreTitleTxt.Text = selected.DisplayName;
                
                if (viewImageBtn != null)
                {
                    bool showButton = false;
                    AfpNode bocNode = null;

                    if (selected.Name == "BOC") bocNode = selected;
                    else if (selected.Name == "OCD") bocNode = selected.Parent;
                    else if (selected.Name == "EOC")
                    {
                        if (selected.Parent != null)
                        {
                            for (int i = selected.Parent.Children.IndexOf(selected) - 1; i >= 0; i--)
                            {
                                if (selected.Parent.Children[i].Name == "BOC")
                                {
                                    bocNode = selected.Parent.Children[i];
                                    break;
                                }
                            }
                        }
                    }

                    if (bocNode != null)
                    {
                        showButton = IsImageContainer(bocNode);
                    }

                    if (showButton)
                    {
                        viewImageBtn.Visibility = Visibility.Visible;
                        viewImageBtn.Tag = selected;
                    }
                    else
                    {
                        viewImageBtn.Visibility = Visibility.Collapsed;
                        viewImageBtn.Tag = null;
                    }
                }
                
                if (exploreDocLinkBlock != null && exploreLink != null)
                {
                    if (!string.IsNullOrEmpty(selected.Url))
                    {
                        if (Uri.TryCreate(selected.Url, UriKind.Absolute, out Uri resultUri))
                        {
                            exploreLink.NavigateUri = resultUri;
                            exploreDocLinkBlock.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            exploreDocLinkBlock.Visibility = Visibility.Collapsed;
                        }
                    }
                    else
                    {
                        exploreDocLinkBlock.Visibility = Visibility.Collapsed;
                    }
                }

                if (exploreOffsetTxt != null) exploreOffsetTxt.Text = $"Offset: 0x{selected.Offset:X8}";
                if (exploreLengthTxt != null) exploreLengthTxt.Text = $"Length: {selected.Length}";

                // Context info for Hexdump Header
                if (hexdumpHeaderTxt != null)
                {
                    hexdumpHeaderTxt.Text = $"SF @ 0x{selected.Offset:X8} Id=0x{selected.Id} Name={selected.Name} - {selected.FriendlyName}";
                }

                try
                {
                    var dataPayload = await Task.Run(() =>
                    {
                        using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                        {
                            reader.BaseStream.Seek(selected.Offset + 9, SeekOrigin.Begin);
                            return reader.ReadBytes(selected.Length);
                        }
                    });

                    // Hexdump Generation
                    string hexdumpText = await Task.Run(() => FormatHexDump(dataPayload, selected.Offset + 9));
                    if (hexdumpTxt != null)
                    {
                        hexdumpTxt.Text = hexdumpText;
                    }

                    // Extract layout data and push to grid
                    string acronym = AfpConstants.GetSfInfo(selected.Id).Acronym;
                    if (acronym != "UNKNOWN" && acronym != "UNK")
                    {
                        var def = MappingService.Instance.GetDefinition(acronym);
                        if (def != null && def.Layout != null && def.Layout.Count > 0)
                        {
                            var parsedData = await Task.Run(() => MappingService.Instance.ParseSfData(acronym, dataPayload, ebcdic, selected.RepeatingGroupLength));
                            if (exploreListView != null) exploreListView.ItemsSource = parsedData.LayoutRows;
                            var TripletsItemsControl = this.FindName("TripletsItemsControl") as ItemsControl;
                            if (TripletsItemsControl != null) TripletsItemsControl.ItemsSource = parsedData.Triplets;
                            
                            if (exploreListView != null) exploreListView.Visibility = Visibility.Visible;
                            if (exploreFallbackTxt != null) exploreFallbackTxt.Visibility = Visibility.Collapsed;

                            var exploreScrollViewer = this.FindName("ExploreScrollViewer") as ScrollViewer;
                            if (exploreScrollViewer != null) exploreScrollViewer.ScrollToTop();
                        }
                        else
                        {
                            if (exploreListView != null) exploreListView.Visibility = Visibility.Collapsed;
                            var TripletsItemsControl = this.FindName("TripletsItemsControl") as ItemsControl;
                            if (TripletsItemsControl != null) TripletsItemsControl.ItemsSource = null;
                            if (exploreFallbackTxt != null) 
                            {
                                exploreFallbackTxt.Visibility = Visibility.Visible;
                                exploreFallbackTxt.Text = $"Definition for {acronym} found, but no detailed layout is mapped yet.";
                            }
                        }
                    }
                    else
                    {
                        if (exploreListView != null) exploreListView.Visibility = Visibility.Collapsed;
                        var TripletsItemsControl = this.FindName("TripletsItemsControl") as ItemsControl;
                        if (TripletsItemsControl != null) TripletsItemsControl.ItemsSource = null;
                        if (exploreFallbackTxt != null) 
                        {
                            exploreFallbackTxt.Visibility = Visibility.Visible;
                            exploreFallbackTxt.Text = $"No definition mapping exists matching Hex ID {selected.Id}.";
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (exploreFallbackTxt != null) 
                    {
                        exploreFallbackTxt.Visibility = Visibility.Visible;
                        exploreFallbackTxt.Text = "Error reading data: " + ex.Message;
                    }
                }
                finally
                {
                    if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Collapsed;
                }
            }
        }

        private bool IsImageContainer(AfpNode bocNode)
        {
            if (bocNode == null || bocNode.Name != "BOC" || string.IsNullOrEmpty(_currentFilePath))
                return false;

            try
            {
                string detectedType = "Unknown Object";
                using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                {
                    reader.BaseStream.Seek(bocNode.Offset + 9, SeekOrigin.Begin);
                    byte[] bocPayload = reader.ReadBytes(bocNode.Length);

                    for (int i = 0; i <= bocPayload.Length - 4; i++)
                    {
                        if (bocPayload[i] == 0x06 && bocPayload[i+1] == 0x07 && bocPayload[i+2] == 0x2B && bocPayload[i+3] == 0x12)
                        {
                            byte[] oidBytes = new byte[16];
                            int tripletRemaining = bocPayload.Length - i;
                            int extractLen = Math.Min(16, tripletRemaining);
                            Array.Copy(bocPayload, i, oidBytes, 0, extractLen);
                            string oidStr = BitConverter.ToString(oidBytes).Replace("-", "");

                            var objectTypes = new Dictionary<string, string>
                            {
                                ["06072B12000401010500000000000000"] = "IOCA FS10",
                                ["06072B12000401010B00000000000000"] = "IOCA FS11",
                                ["06072B12000401010C00000000000000"] = "IOCA FS45",
                                ["06072B12000401010D00000000000000"] = "EPS",
                                ["06072B12000401010E00000000000000"] = "TIFF",
                                ["06072B12000401010F00000000000000"] = "COM Set-up",
                                ["06072B12000401011000000000000000"] = "Tape Label Set-up",
                                ["06072B12000401011100000000000000"] = "DIB, Windows Version",
                                ["06072B12000401011200000000000000"] = "DIB, OS/2 PM Version",
                                ["06072B12000401011300000000000000"] = "PCX",
                                ["06072B12000401011400000000000000"] = "Color Mapping Table (CMT)",
                                ["06072B12000401011600000000000000"] = "GIF",
                                ["06072B12000401011700000000000000"] = "AFPC JPEG Subset",
                                ["06072B12000401011800000000000000"] = "AnaStak Control Record",
                                ["06072B12000401011900000000000000"] = "PDF Single-page Object",
                                ["06072B12000401011A00000000000000"] = "PDF Resource Object",
                                ["06072B12000401012200000000000000"] = "PCL Page Object",
                                ["06072B12000401012D00000000000000"] = "IOCA FS42",
                                ["06072B12000401012E00000000000000"] = "Resident Color Profile",
                                ["06072B12000401012F00000000000000"] = "IOCA Tile Resource",
                                ["06072B12000401013000000000000000"] = "EPS with Transparency",
                                ["06072B12000401013100000000000000"] = "PDF with Transparency",
                                ["06072B12000401013300000000000000"] = "TrueType/OpenType Font",
                                ["06072B12000401013500000000000000"] = "TrueType/OpenType Font Collection",
                                ["06072B12000401013600000000000000"] = "Resource Access Table",
                                ["06072B12000401013700000000000000"] = "IOCA FS40",
                                ["06072B12000401013800000000000000"] = "UP3i Print Data",
                                ["06072B12000401013900000000000000"] = "Color Management Resource (CMR)",
                                ["06072B12000401013A00000000000000"] = "JPEG2000 (JP2) File Format",
                                ["06072B12000401013C00000000000000"] = "TIFF without Transparency",
                                ["06072B12000401013D00000000000000"] = "TIFF Multiple Image File",
                                ["06072B12000401013E00000000000000"] = "TIFF Multiple Image - without Transparency - File",
                                ["06072B12000401013F00000000000000"] = "PDF Multiple Page File",
                                ["06072B12000401014000000000000000"] = "PDF Multiple Page - with Transparency - File",
                                ["06072B12000401014100000000000000"] = "AFPC PNG Subset",
                                ["06072B12000401014200000000000000"] = "AFPC TIFF Subset",
                                ["06072B12000401014300000000000000"] = "Metadata Object",
                                ["06072B12000401014400000000000000"] = "AFPC SVG Subset",
                                ["06072B12000401014500000000000000"] = "Non-OCA Resource Object",
                                ["06072B12000401014600000000000000"] = "IOCA FS48",
                                ["06072B12000401014700000000000000"] = "IOCA FS14"
                            };

                            if (objectTypes.TryGetValue(oidStr, out string typeName))
                            {
                                detectedType = typeName;
                            }
                            break; 
                        }
                    }
                }

                bool isSupportedImage = detectedType.Contains("TIFF") || 
                                        detectedType.Contains("GIF") || 
                                        detectedType.Contains("JPEG") || 
                                        detectedType.Contains("PNG") || 
                                        detectedType.Contains("DIB") || 
                                        detectedType.Contains("PCX");

                return isSupportedImage;
            }
            catch
            {
                return false;
            }
        }

        private void TripletHyperlink_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Hyperlink link && link.Tag is AFPEngineer_Explorer.Models.ParsedTriplet triplet)
            {
                var win = new TripletJsonWindow(triplet.Heading, triplet.RawJson, triplet.Url);
                win.Owner = this;
                win.ShowDialog();
            }
        }

        private void ViewImage_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn?.Tag is AfpNode selectedNode && !string.IsNullOrEmpty(_currentFilePath))
            {
                AfpNode bocNode = null;

                if (selectedNode.Name == "BOC") bocNode = selectedNode;
                else if (selectedNode.Name == "OCD") bocNode = selectedNode.Parent;
                else if (selectedNode.Name == "EOC")
                {
                    // EOC is a sibling of BOC, find the prior BOC sibling. But currently EOC has the same parent as BOC.
                    if (selectedNode.Parent != null)
                    {
                        for (int i = selectedNode.Parent.Children.IndexOf(selectedNode) - 1; i >= 0; i--)
                        {
                            if (selectedNode.Parent.Children[i].Name == "BOC")
                            {
                                bocNode = selectedNode.Parent.Children[i];
                                break;
                            }
                        }
                    }
                }

                if (bocNode == null || bocNode.Name != "BOC")
                {
                    MessageBox.Show("Could not locate the Begin Object Container (BOC) for this element.");
                    return;
                }

                try
                {
                    // Inspect the BOC payload for Object Type OID
                    string detectedType = "Unknown Object";
                    using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                    {
                        reader.BaseStream.Seek(bocNode.Offset + 9, SeekOrigin.Begin);
                        byte[] bocPayload = reader.ReadBytes(bocNode.Length);

                        // Look for standard MO:DCA registry prefix: 06 07 2B 12
                        for (int i = 0; i <= bocPayload.Length - 4; i++)
                        {
                            if (bocPayload[i] == 0x06 && bocPayload[i+1] == 0x07 && bocPayload[i+2] == 0x2B && bocPayload[i+3] == 0x12)
                            {
                                byte[] oidBytes = new byte[16];
                                // Some OIDs might be shorter than 16 bytes depending on the triplet length, we pad with 0s per rules
                                int tripletRemaining = bocPayload.Length - i;
                                // We know it's inside a triplet, but scanning raw is easiest
                                // Just grab up to 16 bytes. MO:DCA OIDs have a specific length often given, but user dictionary assumes padded 16 byte strings.
                                // We'll just grab the length mentioned in dictionary or pad it manually based on what's available
                                int extractLen = Math.Min(16, tripletRemaining);
                                Array.Copy(bocPayload, i, oidBytes, 0, extractLen);
                                string oidStr = BitConverter.ToString(oidBytes).Replace("-", "");

                                var objectTypes = new Dictionary<string, string>
                                {
                                    ["06072B12000401010500000000000000"] = "IOCA FS10",
                                    ["06072B12000401010B00000000000000"] = "IOCA FS11",
                                    ["06072B12000401010C00000000000000"] = "IOCA FS45",
                                    ["06072B12000401010D00000000000000"] = "EPS",
                                    ["06072B12000401010E00000000000000"] = "TIFF",
                                    ["06072B12000401010F00000000000000"] = "COM Set-up",
                                    ["06072B12000401011000000000000000"] = "Tape Label Set-up",
                                    ["06072B12000401011100000000000000"] = "DIB, Windows Version",
                                    ["06072B12000401011200000000000000"] = "DIB, OS/2 PM Version",
                                    ["06072B12000401011300000000000000"] = "PCX",
                                    ["06072B12000401011400000000000000"] = "Color Mapping Table (CMT)",
                                    ["06072B12000401011600000000000000"] = "GIF",
                                    ["06072B12000401011700000000000000"] = "AFPC JPEG Subset",
                                    ["06072B12000401011800000000000000"] = "AnaStak Control Record",
                                    ["06072B12000401011900000000000000"] = "PDF Single-page Object",
                                    ["06072B12000401011A00000000000000"] = "PDF Resource Object",
                                    ["06072B12000401012200000000000000"] = "PCL Page Object",
                                    ["06072B12000401012D00000000000000"] = "IOCA FS42",
                                    ["06072B12000401012E00000000000000"] = "Resident Color Profile",
                                    ["06072B12000401012F00000000000000"] = "IOCA Tile Resource",
                                    ["06072B12000401013000000000000000"] = "EPS with Transparency",
                                    ["06072B12000401013100000000000000"] = "PDF with Transparency",
                                    ["06072B12000401013300000000000000"] = "TrueType/OpenType Font",
                                    ["06072B12000401013500000000000000"] = "TrueType/OpenType Font Collection",
                                    ["06072B12000401013600000000000000"] = "Resource Access Table",
                                    ["06072B12000401013700000000000000"] = "IOCA FS40",
                                    ["06072B12000401013800000000000000"] = "UP3i Print Data",
                                    ["06072B12000401013900000000000000"] = "Color Management Resource (CMR)",
                                    ["06072B12000401013A00000000000000"] = "JPEG2000 (JP2) File Format",
                                    ["06072B12000401013C00000000000000"] = "TIFF without Transparency",
                                    ["06072B12000401013D00000000000000"] = "TIFF Multiple Image File",
                                    ["06072B12000401013E00000000000000"] = "TIFF Multiple Image - without Transparency - File",
                                    ["06072B12000401013F00000000000000"] = "PDF Multiple Page File",
                                    ["06072B12000401014000000000000000"] = "PDF Multiple Page - with Transparency - File",
                                    ["06072B12000401014100000000000000"] = "AFPC PNG Subset",
                                    ["06072B12000401014200000000000000"] = "AFPC TIFF Subset",
                                    ["06072B12000401014300000000000000"] = "Metadata Object",
                                    ["06072B12000401014400000000000000"] = "AFPC SVG Subset",
                                    ["06072B12000401014500000000000000"] = "Non-OCA Resource Object",
                                    ["06072B12000401014600000000000000"] = "IOCA FS48",
                                    ["06072B12000401014700000000000000"] = "IOCA FS14"
                                };

                                if (objectTypes.TryGetValue(oidStr, out string typeName))
                                {
                                    detectedType = typeName;
                                }
                                break; // found the primary OID
                            }
                        }
                    }

                    // Check if the detected type is supported by standard WPF Image control
                    bool isSupportedImage = detectedType.Contains("TIFF") || 
                                            detectedType.Contains("GIF") || 
                                            detectedType.Contains("JPEG") || 
                                            detectedType.Contains("PNG") || 
                                            detectedType.Contains("DIB") || 
                                            detectedType.Contains("PCX");

                    if (!isSupportedImage && detectedType != "Unknown Object")
                    {
                        var imageErrorTxt = this.FindName("ImageErrorTxt") as TextBlock;
                        var containerImage = this.FindName("ContainerImage") as Image;
                        var viewerTab = this.FindName("ImageViewerTab") as TabItem;

                        if (containerImage != null) containerImage.Visibility = Visibility.Collapsed;
                        if (imageErrorTxt != null)
                        {
                            imageErrorTxt.Text = $"This Object Container holds a(n) '{detectedType}'. This format cannot be rendered natively as a standard image.";
                            imageErrorTxt.Visibility = Visibility.Visible;
                        }
                        if (viewerTab != null)
                        {
                            viewerTab.Visibility = Visibility.Visible;
                            viewerTab.IsSelected = true;
                        }
                        return; // Exit early completely so it doesn't even try to crash reading bytes.
                    }

                    AfpNode obdNode = null;
                    if (bocNode.Parent != null)
                    {
                        foreach (var sibling in bocNode.Parent.Children)
                        {
                            if (sibling.Name == "OBD" && sibling.Offset < bocNode.Offset)
                            {
                                obdNode = sibling;
                            }
                        }
                    }

                    double? targetWidth = null;
                    double? targetHeight = null;

                    var ms = new MemoryStream();
                    using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                    {
                        if (obdNode != null)
                        {
                            reader.BaseStream.Seek(obdNode.Offset + 9, SeekOrigin.Begin);
                            byte[] obdData = reader.ReadBytes(obdNode.Length);
                            
                            int idx = 0;
                            double xUnits = 14400, yUnits = 14400; // Defaults
                            double baseMultiplierX = 10, baseMultiplierY = 10;
                            int xSize = 0, ySize = 0;
                            bool found4B = false, found4C = false;

                            while (idx < obdData.Length)
                            {
                                int tLen = obdData[idx];
                                if (tLen < 2 || idx + tLen > obdData.Length) break;
                                byte tId = obdData[idx + 1];

                                if (tId == 0x4B && tLen >= 8) // Measurement
                                {
                                    byte xBase = obdData[idx + 2];
                                    byte yBase = obdData[idx + 3];
                                    baseMultiplierX = xBase == 0x01 ? 3.93701 : 10.0;
                                    baseMultiplierY = yBase == 0x01 ? 3.93701 : 10.0;
                                    xUnits = (obdData[idx + 4] << 8) | obdData[idx + 5];
                                    yUnits = (obdData[idx + 6] << 8) | obdData[idx + 7];
                                    found4B = true;
                                }
                                else if (tId == 0x4C && tLen >= 9) // Size
                                {
                                    xSize = (obdData[idx + 3] << 16) | (obdData[idx + 4] << 8) | obdData[idx + 5];
                                    ySize = (obdData[idx + 6] << 16) | (obdData[idx + 7] << 8) | obdData[idx + 8];
                                    found4C = true;
                                }
                                idx += tLen;
                            }

                            if (found4C && xUnits > 0 && yUnits > 0)
                            {
                                targetWidth = (xSize * baseMultiplierX / xUnits) * 96.0;
                                targetHeight = (ySize * baseMultiplierY / yUnits) * 96.0;
                            }
                        }

                        // Collect all OCD children
                        foreach (var child in bocNode.Children)
                        {
                            if (child.Name == "OCD")
                            {
                                reader.BaseStream.Seek(child.Offset + 9, SeekOrigin.Begin);
                                byte[] ocdData = reader.ReadBytes(child.Length);
                                ms.Write(ocdData, 0, ocdData.Length);
                            }
                        }
                    }

                    ms.Position = 0;
                    if (ms.Length > 0)
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.StreamSource = ms;
                        bitmap.EndInit();

                        var imageErrorTxt = this.FindName("ImageErrorTxt") as TextBlock;
                        var containerImage = this.FindName("ContainerImage") as Image;
                        var viewerTab = this.FindName("ImageViewerTab") as TabItem;

                        if (containerImage != null)
                        {
                            containerImage.Source = bitmap;
                            containerImage.Visibility = Visibility.Visible;
                            if (targetWidth.HasValue && targetHeight.HasValue)
                            {
                                containerImage.Width = targetWidth.Value;
                                containerImage.Height = targetHeight.Value;
                                containerImage.Margin = new Thickness(5);
                            }
                            else
                            {
                                containerImage.Width = double.NaN;
                                containerImage.Height = double.NaN;
                                containerImage.Margin = new Thickness(10);
                            }
                        }

                        if (imageErrorTxt != null) imageErrorTxt.Visibility = Visibility.Collapsed;
                        if (viewerTab != null)
                        {
                            viewerTab.Visibility = Visibility.Visible;
                            viewerTab.IsSelected = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No Object Container Data (OCD) fields found within this container.");
                    }
                }
                catch (Exception ex)
                {
                    var imageErrorTxt = this.FindName("ImageErrorTxt") as TextBlock;
                    var containerImage = this.FindName("ContainerImage") as Image;
                    
                    if (containerImage != null) containerImage.Visibility = Visibility.Collapsed;

                    if (imageErrorTxt != null)
                    {
                        if (ex is NotSupportedException)
                            imageErrorTxt.Text = "Could not render the image. This object container likely holds a non-image format natively unsupported by WPF (e.g., PDF, CMR, or EPS).";
                        else
                            imageErrorTxt.Text = $"Could not render the image. Is it a valid supported format? Error: {ex.Message}";
                        
                        imageErrorTxt.Visibility = Visibility.Visible;
                    }
                    var viewerTab = this.FindName("ImageViewerTab") as TabItem;
                    if (viewerTab != null)
                    {
                        viewerTab.Visibility = Visibility.Visible;
                        viewerTab.IsSelected = true;
                    }
                }
            }
        }

        private void ExplorePrevBtn_Click(object sender, RoutedEventArgs e)
        {
            var tree = this.FindName("AfpTreeView") as TreeView;
            if (tree?.SelectedItem is AfpNode selectedNode)
            {
                int currentIndex = _flatNodeList.IndexOf(selectedNode);
                if (currentIndex > 0)
                {
                    var prevNode = _flatNodeList[currentIndex - 1];
                    prevNode.IsSelected = true;
                    BringNodeIntoView(prevNode, tree);
                }
            }
        }

        private void ExploreNextBtn_Click(object sender, RoutedEventArgs e)
        {
            var tree = this.FindName("AfpTreeView") as TreeView;
            if (tree?.SelectedItem is AfpNode selectedNode)
            {
                int currentIndex = _flatNodeList.IndexOf(selectedNode);
                if (currentIndex >= 0 && currentIndex < _flatNodeList.Count - 1)
                {
                    var nextNode = _flatNodeList[currentIndex + 1];
                    nextNode.IsSelected = true;
                    BringNodeIntoView(nextNode, tree);
                }
            }
        }

        // --- Document Display Handlers ---
        private int _currentDocPage = 1;
        private int _totalDocPages = 1;
        private double _docZoomLevel = 1.0;
        private List<AfpNode> _pageNodes = new List<AfpNode>();

        private void DocPrevPageBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_currentDocPage > 1)
            {
                _currentDocPage--;
                UpdateDocumentDisplay();
            }
        }

        private void DocNextPageBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_currentDocPage < _totalDocPages)
            {
                _currentDocPage++;
                UpdateDocumentDisplay();
            }
        }

        private void DocZoomOutBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_docZoomLevel > 0.25)
            {
                _docZoomLevel -= 0.25;
                ApplyDocumentZoom();
            }
        }

        private void DocZoomInBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_docZoomLevel < 4.0)
            {
                _docZoomLevel += 0.25;
                ApplyDocumentZoom();
            }
        }

        private async void AskAIBtn_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var node = btn?.Tag as AfpNode;
            var aiResponseTxt = this.FindName("AiResponseTxt") as TextBox;
            var exploreListView = this.FindName("ExploreListView") as ListView;
            
            if (node == null || aiResponseTxt == null || exploreListView == null) return;

            var properties = exploreListView.ItemsSource as IEnumerable<dynamic>;
            if (properties == null) return;

            btn.IsEnabled = false;
            btn.Content = "✨ Thinking...";
            aiResponseTxt.Visibility = Visibility.Visible;
            aiResponseTxt.Text = "Generating explanation...\n";

            try
            {
                string githubToken = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
                if (string.IsNullOrEmpty(githubToken))
                {
                    aiResponseTxt.Text = "Please set the GITHUB_TOKEN environment variable. You can set it globally or start the app from a terminal where it's set.";
                    return;
                }
                
                var client = new OpenAI.OpenAIClient(new ApiKeyCredential(githubToken), new OpenAI.OpenAIClientOptions { Endpoint = new Uri("https://models.inference.ai.azure.com") });
                var chatClient = client.GetChatClient("gpt-4o-mini"); // Default fast model

                string propertiesContext = "";
                foreach (var prop in properties)
                {
                    propertiesContext += $"- {prop.DisplayName}: {prop.Value}\n";
                }

                string prompt = $@"
You are an expert in the IBM AFP (Advanced Function Presentation) Datastream specification.
I am analyzing the '{node.Name}' ({node.DisplayName}) structured field. 
Here are the parsed properties for this specific node:
{propertiesContext}

In 3 concise bullet points, explain what this specific field acts as in the datastream, and highlight if any of the property values shown seem irregular or important.";

                var responseStream = chatClient.CompleteChatStreamingAsync(new ChatMessage[]
                {
                    new SystemChatMessage("You are a helpful software engineering assistant specializing in print streams."),
                    new UserChatMessage(prompt)
                });

                aiResponseTxt.Text = "";
                await foreach (var update in responseStream)
                {
                    if (update.ContentUpdate != null)
                    {
                        foreach (var part in update.ContentUpdate)
                        {
                            aiResponseTxt.Text += part.Text;
                            aiResponseTxt.ScrollToEnd(); 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                aiResponseTxt.Text = $"Error getting AI response: {ex.Message}";
            }
            finally
            {
                btn.IsEnabled = true;
                btn.Content = "✨ Explain Field with AI";
            }
        }

        private void ApplyDocumentZoom()
        {
            var zoomTxt = this.FindName("DocZoomTxt") as TextBlock;
            if (zoomTxt != null)
                zoomTxt.Text = $"{(_docZoomLevel * 100):0}%";

            var docCanvas = this.FindName("DocumentCanvas") as Canvas;
            if (docCanvas != null)
            {
                docCanvas.LayoutTransform = new ScaleTransform(_docZoomLevel, _docZoomLevel);
            }
        }

        private async void UpdateDocumentDisplay()
        {
            var pageInfoTxt = this.FindName("DocPageInfoTxt") as TextBlock;
            if (pageInfoTxt != null)
                pageInfoTxt.Text = $"Page {_currentDocPage} of {_totalDocPages}";

            var docCanvas = this.FindName("DocumentCanvas") as Canvas;
            if (docCanvas == null) return;
            
            docCanvas.Children.Clear();

            if (_pageNodes == null || _pageNodes.Count == 0 || _currentDocPage > _pageNodes.Count) return;

            var currentPageNode = _pageNodes[_currentDocPage - 1];

            // Default fallback size (8.5 x 11 inches at 96 DPI)
            double canvasWidth = 816;
            double canvasHeight = 1056;

            // Track state for text positioning
            double currentX = 0;
            double currentY = 0;

            // Default 14400 units per 10 inches (1440 dpi) -> 1440 units per inch
            double resX = 1440.0;
            double resY = 1440.0;

            // Find PGD child for dimensions
            AfpNode pgdNode = null;
            void FindPGD(AfpNode node)
            {
                if (node.Name == "PGD") pgdNode = node;
                else foreach (var c in node.Children) FindPGD(c);
            }
            FindPGD(currentPageNode);

            if (pgdNode != null && !string.IsNullOrEmpty(_currentFilePath))
            {
                try
                {
                    using (var reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                    {
                        reader.BaseStream.Seek(pgdNode.Offset + 9, SeekOrigin.Begin); // skip 9 byte header
                        byte[] pgdData = reader.ReadBytes(pgdNode.Length);

                        if (pgdData.Length >= 12)
                        {
                            byte xBase = pgdData[0];
                            byte yBase = pgdData[1];
                            int xUnits = (pgdData[2] << 8) | pgdData[3];
                            int yUnits = (pgdData[4] << 8) | pgdData[5];
                            int xSize = (pgdData[6] << 16) | (pgdData[7] << 8) | pgdData[8];
                            int ySize = (pgdData[9] << 16) | (pgdData[10] << 8) | pgdData[11];

                            double baseMultiplierX = xBase == 0x01 ? 3.93701 : 10.0;
                            double baseMultiplierY = yBase == 0x01 ? 3.93701 : 10.0;

                            if (xUnits > 0 && yUnits > 0)
                            {
                                resX = xUnits / baseMultiplierX; // Units per inch
                                resY = yUnits / baseMultiplierY; // Units per inch

                                canvasWidth = (xSize / resX) * 96.0;
                                canvasHeight = (ySize / resY) * 96.0;
                            }
                        }
                    }
                }
                catch { } // Ignore errors, keep default size
            }

            docCanvas.Width = canvasWidth;
            docCanvas.Height = canvasHeight;

            // Find PTD (Presentation Text Descriptor) which defines the units for PTOCA movements
            // PTD is usually inside the Active Environment Group (BAG)
            double textResX = resX;
            double textResY = resY;
            AfpNode ptdNode = null;
            void FindPTD(AfpNode node)
            {
                if (node.Name == "PTD" || node.Id == "D3B19B") ptdNode = node;
                if (ptdNode == null)
                {
                    foreach (var child in node.Children) FindPTD(child);
                }
            }
            FindPTD(currentPageNode);

            if (ptdNode != null && !string.IsNullOrEmpty(_currentFilePath))
            {
                try
                {
                    using (var reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                    {
                        reader.BaseStream.Seek(ptdNode.Offset + 9, SeekOrigin.Begin);
                        byte[] ptdData = reader.ReadBytes(ptdNode.Length);

                        if (ptdData.Length >= 6)
                        {
                            byte xpBase = ptdData[0];
                            byte ypBase = ptdData[1];
                            int xpUnits = (ptdData[2] << 8) | ptdData[3];
                            int ypUnits = (ptdData[4] << 8) | ptdData[5];

                            double baseMultiplierX = xpBase == 0x01 ? 3.93701 : 10.0;
                            double baseMultiplierY = ypBase == 0x01 ? 3.93701 : 10.0;

                            if (xpUnits > 0 && ypUnits > 0)
                            {
                                textResX = xpUnits / baseMultiplierX;
                                textResY = ypUnits / baseMultiplierY;
                            }
                        }
                    }
                }
                catch { } // fallback to PGD
            }

            int ptxPayloadsParsed = 0;
            int textBlocksAdded = 0;

            void ProcessNodeForPTX(AfpNode currentNode)
            {
                if (currentNode.Name == "PTX")
                {
                    ptxPayloadsParsed++;
                    try
                    {
                        using (var reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                        {
                            reader.BaseStream.Seek(currentNode.Offset + 9, SeekOrigin.Begin);
                            byte[] payload = reader.ReadBytes(currentNode.Length);
                            int pos = 0;

                            while (pos < payload.Length)
                            {
                                // Skip PTOCA escape sequence if present
                                if (pos + 1 < payload.Length && payload[pos] == 0x2B && payload[pos + 1] == 0xD3)
                                {
                                    pos += 2;
                                    continue;
                                }

                                byte lenByte = payload[pos];
                                int len = lenByte == 0 ? 256 : (int)lenByte;
                                
                                if (len < 2 || pos + len > payload.Length)
                                {
                                    // If we somehow lost sync and hit plain text, we would misinterpret it. 
                                    // For now, assuming pure control sequence format as often seen.
                                    break;
                                }

                                byte csCode = payload[pos + 1];
                                int dataStart = pos + 2;
                                int dataLen = len - 2;

                                if (csCode == 0xD3 || csCode == 0xD2) // AMB
                                {
                                    int val = (payload[dataStart] << 8) | payload[dataStart + 1];
                                    currentY = (val / textResY) * 96.0;
                                }
                                else if (csCode == 0xC7 || csCode == 0xC6) // AMI
                                {
                                    int val = (payload[dataStart] << 8) | payload[dataStart + 1];
                                    currentX = (val / textResX) * 96.0;
                                }
                                else if (csCode == 0xD5 || csCode == 0xD4) // RMB
                                {
                                    short val = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                    currentY += (val / textResY) * 96.0;
                                }
                                else if (csCode == 0xC9 || csCode == 0xC8) // RMI
                                {
                                    short val = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                    currentX += (val / textResX) * 96.0;
                                }
                                else if (csCode == 0xDB || csCode == 0xDA) // TRN
                                {
                                    if (dataLen > 0)
                                    {
                                        byte[] trnBytes = new byte[dataLen];
                                        Array.Copy(payload, dataStart, trnBytes, 0, dataLen);
                                        string text = ebcdic.GetString(trnBytes);
                                        
                                        TextBlock tb = new TextBlock
                                        {
                                            Text = text,
                                            FontFamily = new FontFamily("Courier New"),
                                            FontSize = 10,
                                            Foreground = System.Windows.Media.Brushes.Black
                                        };
                                        Canvas.SetLeft(tb, currentX);
                                        Canvas.SetTop(tb, currentY - 8);
                                        docCanvas.Children.Add(tb);
                                        textBlocksAdded++;
                                        
                                        // Some fonts are roughly 6-7 pts wide. But TRN often moves coordinate via RMI/AMI.
                                        currentX += text.Length * 6.0; 
                                    }
                                }
                                else if (csCode == 0xE7 || csCode == 0xE6) // DBR (Draw B-axis Rule)
                                {
                                    if (dataLen >= 2)
                                    {
                                        short ruleLength = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                        double rLength = (ruleLength / textResY) * 96.0;

                                        double rWidth = 1.0; // Default width
                                        if (dataLen >= 5)
                                        {
                                            short wVal1 = (short)((payload[dataStart + 2] << 8) | payload[dataStart + 3]);
                                            byte wVal2 = payload[dataStart + 4];
                                            double fraction = 0;
                                            for (int bit = 0; bit < 8; bit++)
                                            {
                                                if ((wVal2 & (1 << (7 - bit))) != 0)
                                                    fraction += 1.0 / (1 << (bit + 1));
                                            }
                                            rWidth = ((wVal1 + fraction) / textResX) * 96.0;
                                        }

                                        System.Windows.Shapes.Rectangle rect = new System.Windows.Shapes.Rectangle
                                        {
                                            Width = Math.Max(1.0, Math.Abs(rWidth)),
                                            Height = Math.Max(1.0, Math.Abs(rLength)),
                                            Fill = System.Windows.Media.Brushes.Black
                                        };
                                        Canvas.SetLeft(rect, rWidth < 0 ? currentX + rWidth : currentX);
                                        Canvas.SetTop(rect, rLength < 0 ? currentY + rLength : currentY);
                                        docCanvas.Children.Add(rect);
                                    }
                                }
                                else if (csCode == 0xE5 || csCode == 0xE4) // DIR (Draw I-axis Rule)
                                {
                                    if (dataLen >= 2)
                                    {
                                        short ruleLength = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                        double rLength = (ruleLength / textResX) * 96.0;

                                        double rWidth = 1.0; // Default width
                                        if (dataLen >= 5)
                                        {
                                            short wVal1 = (short)((payload[dataStart + 2] << 8) | payload[dataStart + 3]);
                                            byte wVal2 = payload[dataStart + 4];
                                            double fraction = 0;
                                            for (int bit = 0; bit < 8; bit++)
                                            {
                                                if ((wVal2 & (1 << (7 - bit))) != 0)
                                                    fraction += 1.0 / (1 << (bit + 1));
                                            }
                                            rWidth = ((wVal1 + fraction) / textResY) * 96.0;
                                        }

                                        System.Windows.Shapes.Rectangle rect = new System.Windows.Shapes.Rectangle
                                        {
                                            Width = Math.Max(1.0, Math.Abs(rLength)),
                                            Height = Math.Max(1.0, Math.Abs(rWidth)),
                                            Fill = System.Windows.Media.Brushes.Black
                                        };
                                        Canvas.SetLeft(rect, rLength < 0 ? currentX + rLength : currentX);
                                        Canvas.SetTop(rect, rWidth < 0 ? currentY + rWidth : currentY);
                                        docCanvas.Children.Add(rect);
                                    }
                                }

                                pos += len;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error rendering page text PTX: {ex.Message}");
                    }
                }

                foreach (var child in currentNode.Children)
                {
                    ProcessNodeForPTX(child);
                }
            }
            
            async Task ProcessNodeForIOB(AfpNode currentNode)
            {
                if (currentNode.Name == "IOB")
                {
                    try
                    {
                        using (var reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                        {
                            reader.BaseStream.Seek(currentNode.Offset + 9, SeekOrigin.Begin);
                            byte[] payload = reader.ReadBytes(currentNode.Length);
                            
                            if (payload != null && payload.Length >= 16)
                            {
                                // ObjName is at offset 0, length 8
                                string targetName = ebcdic.GetString(payload.Take(8).ToArray());
                                
                                // Try basic coordinates based on mapping (SBIN)
                                int xOset = (payload[10] << 16) | (payload[11] << 8) | payload[12];
                                if ((xOset & 0x800000) != 0) xOset -= 0x1000000; // Sign extend 24-bit
                                
                                int yOset = (payload[13] << 16) | (payload[14] << 8) | payload[15];
                                if ((yOset & 0x800000) != 0) yOset -= 0x1000000; // Sign extend 24-bit
                                
                                double xOffset = (xOset / resX) * 96.0;
                                double yOffset = (yOset / resY) * 96.0;

                                System.Diagnostics.Debug.WriteLine($"Found IOB target '{targetName}' at {xOffset}, {yOffset}");

                                // Find matching BOC node anywhere in the document
                                AfpNode targetBoc = null;
                                foreach (var node in _flatNodeList)
                                {
                                    if (node.Name == "BOC")
                                    {
                                        reader.BaseStream.Seek(node.Offset + 9, SeekOrigin.Begin);
                                        byte[] bocPayload = reader.ReadBytes(Math.Min(8, node.Length));
                                        if (bocPayload.Length >= 8)
                                        {
                                            string bocName = ebcdic.GetString(bocPayload);
                                            if (bocName.Trim() == targetName.Trim())
                                            {
                                                targetBoc = node;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (targetBoc != null)
                                {
                                    // Extract OBD for scaling
                                    AfpNode obdNode = null;
                                    if (targetBoc.Parent != null)
                                    {
                                        foreach (var sibling in targetBoc.Parent.Children)
                                        {
                                            if (sibling.Name == "OBD" && sibling.Offset < targetBoc.Offset)
                                            {
                                                obdNode = sibling;
                                            }
                                        }
                                    }

                                    double? targetWidth = null;
                                    double? targetHeight = null;

                                    if (obdNode != null)
                                    {
                                        reader.BaseStream.Seek(obdNode.Offset + 9, SeekOrigin.Begin);
                                        byte[] obdData = reader.ReadBytes(obdNode.Length);
                                        
                                        int idx = 0;
                                        double xUnits = 14400, yUnits = 14400;
                                        double baseMultiplierX = 10, baseMultiplierY = 10;
                                        int xSize = 0, ySize = 0;
                                        bool found4C = false;
                                        
                                        while (idx < obdData.Length)
                                        {
                                            int tLen = obdData[idx];
                                            if (tLen < 2 || idx + tLen > obdData.Length) break;
                                            byte tId = obdData[idx + 1];
                                            
                                            if (tId == 0x4B && tLen >= 8) // Measurement
                                            {
                                                byte xBase = obdData[idx + 2];
                                                byte yBase = obdData[idx + 3];
                                                baseMultiplierX = xBase == 0x01 ? 3.93701 : 10.0;
                                                baseMultiplierY = yBase == 0x01 ? 3.93701 : 10.0;
                                                xUnits = (obdData[idx + 4] << 8) | obdData[idx + 5];
                                                yUnits = (obdData[idx + 6] << 8) | obdData[idx + 7];
                                            }
                                            else if (tId == 0x4C && tLen >= 9) // Size
                                            {
                                                xSize = (obdData[idx + 3] << 16) | (obdData[idx + 4] << 8) | obdData[idx + 5];
                                                ySize = (obdData[idx + 6] << 16) | (obdData[idx + 7] << 8) | obdData[idx + 8];
                                                found4C = true;
                                            }
                                            idx += tLen;
                                        }

                                        if (found4C && xUnits > 0 && yUnits > 0)
                                        {
                                            targetWidth = (xSize * baseMultiplierX / xUnits) * 96.0;
                                            targetHeight = (ySize * baseMultiplierY / yUnits) * 96.0;
                                        }
                                    }

                                    // Check object type (PDF vs Image)
                                    bool isPdf = false;
                                    reader.BaseStream.Seek(targetBoc.Offset + 9, SeekOrigin.Begin);
                                    byte[] bocFullPayload = reader.ReadBytes(targetBoc.Length);
                                    for (int i = 0; i <= bocFullPayload.Length - 4; i++)
                                    {
                                        if (bocFullPayload[i] == 0x06 && bocFullPayload[i+1] == 0x07 && bocFullPayload[i+2] == 0x2B && bocFullPayload[i+3] == 0x12)
                                        {
                                            byte[] oidBytes = new byte[16];
                                            int extractLen = Math.Min(16, bocFullPayload.Length - i);
                                            Array.Copy(bocFullPayload, i, oidBytes, 0, extractLen);
                                            string oidStr = BitConverter.ToString(oidBytes).Replace("-", "");
                                            
                                            // Known PDF OIDs in MO:DCA
                                            if (oidStr == "06072B12000401011900000000000000" || // Single-page PDF
                                                oidStr == "06072B12000401011A00000000000000" || // PDF Resource
                                                oidStr == "06072B12000401013100000000000000" || // PDF with transparency
                                                oidStr == "06072B12000401013F00000000000000" || // Multiple Page PDF
                                                oidStr == "06072B12000401014000000000000000")   // Multiple Page PDF with transparency
                                            {
                                                isPdf = true;
                                            }
                                            break;
                                        }
                                    }

                                    // Extract OCD data and map to image / pdf
                                    using (var ms = new MemoryStream())
                                    {
                                        foreach (var child in targetBoc.Children)
                                        {
                                            if (child.Name == "OCD")
                                            {
                                                reader.BaseStream.Seek(child.Offset + 9, SeekOrigin.Begin);
                                                byte[] ocdData = reader.ReadBytes(child.Length);
                                                ms.Write(ocdData, 0, ocdData.Length);
                                            }
                                        }
                                        
                                        ms.Position = 0;
                                        if (ms.Length > 0)
                                        {
                                            if (isPdf)
                                            {
                                                string tempPdf = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");
                                                File.WriteAllBytes(tempPdf, ms.ToArray());

                                                bool renderedViaGs = false;
                                                string gsPath = null;
                                                string[] gsDirs = { @"C:\Program Files\gs", @"C:\Program Files (x86)\gs" };
                                                foreach (var dir in gsDirs)
                                                {
                                                    if (Directory.Exists(dir))
                                                    {
                                                        var exes = Directory.GetFiles(dir, "gswin64c.exe", SearchOption.AllDirectories);
                                                        if (exes.Length > 0) { gsPath = exes[0]; break; }
                                                        exes = Directory.GetFiles(dir, "gswin32c.exe", SearchOption.AllDirectories);
                                                        if (exes.Length > 0) { gsPath = exes[0]; break; }
                                                    }
                                                }

                                                if (!string.IsNullOrEmpty(gsPath))
                                                {
                                                    string tempPng = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                                                    try
                                                    {
                                                        await Task.Run(() =>
                                                        {
                                                            var proc = new System.Diagnostics.Process();
                                                            proc.StartInfo.FileName = gsPath;
                                                            proc.StartInfo.Arguments = $"-dQUIET -dSAFER -dBATCH -dNOPAUSE -dNOPROMPT -sDEVICE=png16m -r300 -sOutputFile=\"{tempPng}\" \"{tempPdf}\"";
                                                            proc.StartInfo.UseShellExecute = false;
                                                            proc.StartInfo.CreateNoWindow = true;
                                                            proc.Start();
                                                            proc.WaitForExit();
                                                        });

                                                        if (File.Exists(tempPng))
                                                        {
                                                            var bmp = new BitmapImage();
                                                            bmp.BeginInit();
                                                            bmp.CacheOption = BitmapCacheOption.OnLoad;
                                                            bmp.UriSource = new Uri(tempPng);
                                                            bmp.EndInit();

                                                            var img = new Image { Source = bmp, Stretch = Stretch.Fill };
                                                            if (targetWidth.HasValue && targetHeight.HasValue)
                                                            {
                                                                img.Width = targetWidth.Value;
                                                                img.Height = targetHeight.Value;
                                                            }
                                                            
                                                            Canvas.SetLeft(img, xOffset);
                                                            Canvas.SetTop(img, yOffset);
                                                            docCanvas.Children.Add(img);
                                                            renderedViaGs = true;
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"Failed to process PDF via Ghostscript: {ex.Message}");
                                                    }
                                                }

                                                if (!renderedViaGs)
                                                {
                                                    // The WPF WebBrowser control uses legacy IE which doesn't have a modern PDF renderer built-in
                                                    // So instead of a WebBrowser (which causes a gray box), we render a clickable proxy button.
                                                    var pdfPlaceholder = new Button
                                                    {
                                                        Content = "📄 Missing Ghostscript. Open PDF Object via External App",
                                                        Background = new SolidColorBrush(Color.FromArgb(100, 200, 200, 255)),
                                                        BorderBrush = Brushes.Blue,
                                                        BorderThickness = new Thickness(1),
                                                        Cursor = System.Windows.Input.Cursors.Hand
                                                    };

                                                    if (targetWidth.HasValue && targetHeight.HasValue)
                                                    {
                                                        pdfPlaceholder.Width = targetWidth.Value;
                                                        pdfPlaceholder.Height = targetHeight.Value;
                                                    }

                                                    // Click launches the system's default PDF viewer (Edge, Acrobat, etc)
                                                    pdfPlaceholder.Click += (s, e) => 
                                                    {
                                                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                                        {
                                                            FileName = tempPdf,
                                                            UseShellExecute = true
                                                        });
                                                    };
                                                    
                                                    Canvas.SetLeft(pdfPlaceholder, xOffset);
                                                    Canvas.SetTop(pdfPlaceholder, yOffset);
                                                    docCanvas.Children.Add(pdfPlaceholder);
                                                }
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    var bitmap = new BitmapImage();
                                                    bitmap.BeginInit();
                                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                                    bitmap.StreamSource = ms;
                                                    bitmap.EndInit();

                                                    var img = new Image { Source = bitmap };
                                                    if (targetWidth.HasValue && targetHeight.HasValue)
                                                    {
                                                        img.Width = targetWidth.Value;
                                                        img.Height = targetHeight.Value;
                                                    }
                                                    
                                                    Canvas.SetLeft(img, xOffset);
                                                    Canvas.SetTop(img, yOffset);
                                                    docCanvas.Children.Add(img);
                                                }
                                                catch (Exception imgEx)
                                                {
                                                    System.Diagnostics.Debug.WriteLine($"Failed to render IOB as Image: {imgEx.Message}");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing IOB: {ex.Message}");
                    }
                }

                foreach (var child in currentNode.Children)
                {
                    await ProcessNodeForIOB(child);
                }
            }

            try
            {
                var loadingOverlay = this.FindName("LoadingOverlay") as Grid;
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Visible;
                ProcessNodeForPTX(currentPageNode);
                await ProcessNodeForIOB(currentPageNode);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing page for rendering: {ex.Message}");
            }
            finally
            {
                var loadingOverlay = this.FindName("LoadingOverlay") as Grid;
                if (loadingOverlay != null) loadingOverlay.Visibility = Visibility.Collapsed;
            }
            
            System.Diagnostics.Debug.WriteLine($"Parsed {ptxPayloadsParsed} PTX payloads, added {textBlocksAdded} text blocks to canvas.");
        }
    }
}
