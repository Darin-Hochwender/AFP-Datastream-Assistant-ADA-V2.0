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
using System.Text.Json;

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

    public class ChatMessageModel : System.ComponentModel.INotifyPropertyChanged
    {
        private string _text;
        public string Text
        {
            get => _text;
            set { _text = value; OnPropertyChanged(nameof(Text)); }
        }

        public System.Windows.Media.Brush BackgroundBrush { get; set; }
        public System.Windows.Media.Brush ForegroundBrush { get; set; }
        public HorizontalAlignment Alignment { get; set; }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string prop) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(prop));
    }

    public partial class MainWindow : Window
    {
        // EBCDIC (Code Page 37) is the standard for most AFP files
        private static Encoding ebcdic;
        
        private ObservableCollection<ChatMessageModel> _chatHistoryUI = new ObservableCollection<ChatMessageModel>();
        private List<ChatMessage> _llmChatHistory = new List<ChatMessage>();
        
        public MainWindow()
        {
            // Add this line to unlock EBCDIC support
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            SettingsService.Load();
            try 
            {
                string cp = SettingsService.CurrentSettings.DefaultCodePage ?? "IBM037";
                if (int.TryParse(cp, out int cpInt))
                    ebcdic = Encoding.GetEncoding(cpInt);
                else
                    ebcdic = Encoding.GetEncoding(cp);
            }
            catch 
            {
                ebcdic = Encoding.GetEncoding("IBM037");
            }

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
            
            var chatList = this.FindName("ChatHistoryList") as ItemsControl;
            if (chatList != null)
            {
                chatList.ItemsSource = _chatHistoryUI;
            }
            
            _llmChatHistory.Add(new SystemChatMessage("You are ADA (AFP Datastream Assistant), a specialized IBM AFP (Advanced Function Presentation) Datastream expert. You help users learn about the datastream, MO:DCA structures, and PTOCA/IOCA. Be concise but highly knowledgeable. Answer questions directly using markdown where helpful, but do not use code blocks unless demonstrating hex or structured fields."));
            _chatHistoryUI.Add(new ChatMessageModel { 
                Text = "Hello! I am ADA, your AFP Datastream Assistant. Ask me any questions about the datastream, structured fields, or triplets!", 
                BackgroundBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F8FF")), 
                ForegroundBrush = System.Windows.Media.Brushes.Black, 
                Alignment = HorizontalAlignment.Left 
            });

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

        private void ShowSettings_Click(object sender, RoutedEventArgs e)
        {
            var win = new SettingsWindow();
            if (win.ShowDialog() == true)
            {
                // Re-render if there's a document loaded
                if (!string.IsNullOrEmpty(_currentFilePath))
                {
                    try 
                    {
                        string cp = SettingsService.CurrentSettings.DefaultCodePage ?? "IBM037";
                        if (int.TryParse(cp, out int cpInt))
                            ebcdic = Encoding.GetEncoding(cpInt);
                        else
                            ebcdic = Encoding.GetEncoding(cp);
                    }
                    catch 
                    {
                        ebcdic = Encoding.GetEncoding("IBM037");
                    }
                    UpdateDocumentDisplay();
                }
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

        private async void AiSearchBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchBox.Text))
            {
                MessageBox.Show("Please enter a natural language search query. Example: 'Find the triplet that defines color management'");
                return;
            }

            string githubToken = GetGithubToken();
            if (string.IsNullOrEmpty(githubToken))
            {
                MessageBox.Show("No GitHub token found. Please set your token to use AI Search.");
                return;
            }

            string term = SearchBox.Text;
            var aiSearchBtn = this.FindName("AiSearchBtn") as Button;
            if (aiSearchBtn != null) { aiSearchBtn.IsEnabled = false; aiSearchBtn.Content = "Wait..."; }

            try
            {
                string sfJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "definitions", "structured_fields.json");
                string tripJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "definitions", "triplets.json");
                string ptxJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "definitions", "ptx_control_sequences.json");

                string sfStr = "";
                if (File.Exists(sfJsonPath))
                {
                    try {
                        using (var doc = JsonDocument.Parse(File.ReadAllText(sfJsonPath))) {
                            if (doc.RootElement.TryGetProperty("sfDefinitions", out var defs)) {
                                foreach (var el in defs.EnumerateArray()) {
                                    string id = el.TryGetProperty("sfId", out var prop) ? prop.GetString() : "";
                                    string desc = el.TryGetProperty("description", out var prop2) ? prop2.GetString() : "";
                                    if (!string.IsNullOrEmpty(id)) sfStr += $"{id}: {desc}\n";
                                }
                            }
                        }
                    } catch {}
                }

                string tripStr = "";
                if (File.Exists(tripJsonPath))
                {
                    try {
                        using (var doc = JsonDocument.Parse(File.ReadAllText(tripJsonPath))) {
                            foreach (var el in doc.RootElement.EnumerateArray()) {
                                string id = el.TryGetProperty("id", out var prop) ? prop.GetString() : "";
                                string name = el.TryGetProperty("name", out var prop3) ? prop3.GetString() : "";
                                if (!string.IsNullOrEmpty(id)) tripStr += $"Hex {id}: {name}\n";
                            }
                        }
                    } catch {}
                }

                string ptxStr = "";
                if (File.Exists(ptxJsonPath))
                {
                    try {
                        using (var doc = JsonDocument.Parse(File.ReadAllText(ptxJsonPath))) {
                            if (doc.RootElement.TryGetProperty("controlSequences", out var seqs)) {
                                foreach (var el in seqs.EnumerateArray()) {
                                    string abbrev = el.TryGetProperty("shortId", out var prop) ? prop.GetString() : "";
                                    string name = el.TryGetProperty("name", out var prop2) ? prop2.GetString() : "";
                                    if (!string.IsNullOrEmpty(abbrev)) ptxStr += $"{abbrev}: {name}\n";
                                }
                            }
                        }
                    } catch {}
                }

                var client = new OpenAI.OpenAIClient(new ApiKeyCredential(githubToken), new OpenAI.OpenAIClientOptions { Endpoint = new Uri("https://models.inference.ai.azure.com") });
                var chatClient = client.GetChatClient("gpt-4o-mini");

                string systemPrompt = $@"You are a Natural Language to AFP element mapper.
Your available schema definitions:
<StructuredFields>{sfStr}</StructuredFields>
<Triplets>{tripStr}</Triplets>
<PTX>{ptxStr}</PTX>

When the user gives a natural language query, find the BEST matching single element.
Respond with a concise, helpful explanation of what the element is.
Then, if the finding maps to a specific Structured Field (e.g. page descriptor, presentation text, etc), you MUST include on a new line: `[SEARCH:ACRONYM]` (e.g. `[SEARCH:PGD]`).
If it maps to a triplet or PTX, mention it in the textual explanation, but ALSO try to provide the most likely structured field they exist in (like `[SEARCH:PTX]` for presentation text control sequences).";

                var response = await chatClient.CompleteChatAsync(new ChatMessage[]
                {
                    new SystemChatMessage(systemPrompt),
                    new UserChatMessage(term)
                });

                string answer = response.Value.Content[0].Text;

                string extractedSearch = null;
                var match = Regex.Match(answer, @"\[SEARCH:([A-Z0-9]{3})\]");
                if (match.Success)
                {
                    extractedSearch = match.Groups[1].Value;
                    answer = answer.Replace(match.Value, "").Trim();
                }

                MessageBox.Show(answer, "AI Search Result", MessageBoxButton.OK, MessageBoxImage.Information);

                if (!string.IsNullOrEmpty(extractedSearch))
                {
                    SearchBox.Text = extractedSearch;
                    SearchAction(1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error running AI Search: {ex.Message}");
            }
            finally
            {
                if (aiSearchBtn != null) { aiSearchBtn.IsEnabled = true; aiSearchBtn.Content = "✨ AI Find"; }
            }
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
            var askAiGroupBtn = this.FindName("AskAIGroupBtn") as Button;
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
                
                if (askAiGroupBtn != null)
                {
                    if (selected.Children != null && selected.Children.Count > 0)
                    {
                        askAiGroupBtn.Visibility = Visibility.Visible;
                        askAiGroupBtn.Tag = selected;
                    }
                    else
                    {
                        askAiGroupBtn.Visibility = Visibility.Collapsed;
                        askAiGroupBtn.Tag = null;
                    }
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

                    bool showPdfButton = false;
                    if (bocNode != null)
                    {
                        showButton = IsImageContainer(bocNode);
                        showPdfButton = IsPdfContainer(bocNode);
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

                    var viewPdfBtn = this.FindName("ViewPdfBtn") as Button;
                    if (viewPdfBtn != null)
                    {
                        if (showPdfButton)
                        {
                            viewPdfBtn.Visibility = Visibility.Visible;
                            viewPdfBtn.Tag = selected;
                        }
                        else
                        {
                            viewPdfBtn.Visibility = Visibility.Collapsed;
                            viewPdfBtn.Tag = null;
                        }
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
                            int extractLen = Math.Min(9, tripletRemaining); // Standard OID prefix is 9 bytes
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
                                // Just grab up to 9 bytes. MO:DCA OIDs have a specific length often given, but user dictionary assumes padded 16 byte strings.
                                // We'll just grab the 9 byte prefix and let the 16 byte array 00-pad the remaining bytes.
                                int extractLen = Math.Min(9, tripletRemaining);
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

        private bool IsPdfContainer(AfpNode bocNode)
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
                        // Look for OID prefix
                        if (bocPayload[i] == 0x06 && bocPayload[i+1] == 0x07 && bocPayload[i+2] == 0x2B && bocPayload[i+3] == 0x12)
                        {
                            byte[] oidBytes = new byte[16];
                            int tripletRemaining = bocPayload.Length - i;
                            int extractLen = Math.Min(9, tripletRemaining);
                            Array.Copy(bocPayload, i, oidBytes, 0, extractLen);
                            string oidStr = BitConverter.ToString(oidBytes).Replace("-", "");

                            var objectTypes = new Dictionary<string, string>
                            {
                                ["06072B12000401011900000000000000"] = "PDF Single-page Object",
                                ["06072B12000401011A00000000000000"] = "PDF Resource Object",
                                ["06072B12000401013100000000000000"] = "PDF with Transparency",
                                ["06072B12000401013F00000000000000"] = "PDF Multiple Page File",
                                ["06072B12000401014000000000000000"] = "PDF Multiple Page - with Transparency - File"
                            };

                            if (objectTypes.TryGetValue(oidStr, out string typeName))
                            {
                                detectedType = typeName;
                            }
                            break;
                        }
                    }
                }

                return detectedType.Contains("PDF");
            }
            catch
            {
                return false;
            }
        }

        private async void ViewPdf_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn?.Tag is AfpNode selectedNode && !string.IsNullOrEmpty(_currentFilePath))
            {
                AfpNode bocNode = null;

                if (selectedNode.Name == "BOC") bocNode = selectedNode;
                else if (selectedNode.Name == "OCD") bocNode = selectedNode.Parent;
                else if (selectedNode.Name == "EOC")
                {
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

                if (bocNode == null || bocNode.Name != "BOC") return;

                try
                {
                    var ms = new MemoryStream();
                    using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                    {
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
                        string tempPdfPath = Path.Combine(Path.GetTempPath(), $"extracted_{Guid.NewGuid()}.pdf");
                        await File.WriteAllBytesAsync(tempPdfPath, ms.ToArray());

                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = tempPdfPath,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        MessageBox.Show("No Object Container Data (OCD) fields found to extract PDF.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error extracting or opening PDF: {ex.Message}");
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

        private async void ExtractTleReport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath) || _flatNodeList == null || _flatNodeList.Count == 0)
            {
                MessageBox.Show("Please open an AFP file first.");
                return;
            }

            string downloadFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string fileName = $"TLE_Report_{Path.GetFileNameWithoutExtension(_currentFilePath)}_{DateTime.Now:yyyyMMddHHmmss}.csv";
            string outputPath = Path.Combine(downloadFolder, fileName);

            var progressWindow = new Window()
            {
                Title = "Extracting TLEs",
                Width = 350,
                Height = 120,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Content = new TextBlock() { Text = "Extracting TLEs to CSV, please wait...", Margin = new Thickness(20), FontSize = 14, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
            };
            progressWindow.Show();

            try
            {
                await Task.Run(() => GenerateTleCsv(outputPath));
                MessageBox.Show($"TLE Report successfully saved to:\n{outputPath}", "Extraction Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error extracting TLEs: {ex.Message}", "Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressWindow.Close();
            }
        }

        private void GenerateTleCsv(string outputPath)
        {
            string[] targetFields = new[] {
                "AccountId", "AddressLine1", "AddressLine2", "AddressLine3", "AddressLine4",
                "AddressLine5", "AddressLine6", "CommId", "DocumentDescription", "DocumentSubType",
                "IDCardCount", "Insert1", "Insert2", "Insert3", "Insert4", "Insert5", "Insert6",
                "Insert7", "Insert8", "Insert9", "Insert10", "International", "LetterDate",
                "OME", "RecipientId", "ReconId", "ZipCd"
            };

            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", targetFields));

            using (BinaryReader reader = new BinaryReader(File.OpenRead(_currentFilePath)))
            {
                Dictionary<string, string> currentGroupValues = null;
                bool inGroup = false;

                foreach (var node in _flatNodeList)
                {
                    if (node.Name == "BNG")
                    {
                        if (inGroup && currentGroupValues != null && currentGroupValues.Count > 0)
                        {
                            var row = targetFields.Select(f => currentGroupValues.ContainsKey(f) ? $"\"{currentGroupValues[f].Replace("\"", "\"\"")}\"" : "");
                            sb.AppendLine(string.Join(",", row));
                        }

                        currentGroupValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        inGroup = true;
                    }
                    else if (inGroup && node.Name == "TLE")
                    {
                        reader.BaseStream.Seek(node.Offset + 9, SeekOrigin.Begin);
                        byte[] tlePayload = reader.ReadBytes(node.Length);
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
                        
                        if (!string.IsNullOrEmpty(attrName))
                        {
                            currentGroupValues[attrName] = attrValue;
                        }
                    }
                    else if (inGroup && node.Name == "ENG")
                    {
                        if (currentGroupValues != null && currentGroupValues.Count > 0)
                        {
                            var row = targetFields.Select(f => currentGroupValues.ContainsKey(f) ? $"\"{currentGroupValues[f].Replace("\"", "\"\"")}\"" : "");
                            sb.AppendLine(string.Join(",", row));
                        }
                        currentGroupValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        inGroup = false;
                    }
                }
                
                if (inGroup && currentGroupValues != null && currentGroupValues.Count > 0)
                {
                    var row = targetFields.Select(f => currentGroupValues.ContainsKey(f) ? $"\"{currentGroupValues[f].Replace("\"", "\"\"")}\"" : "");
                    sb.AppendLine(string.Join(",", row));
                }
            }

            File.WriteAllText(outputPath, sb.ToString());
        }

        private void ViewHelpTopics_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string helpFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "definitions", "help_topics.json");
                if (!File.Exists(helpFilePath))
                {
                    MessageBox.Show("Help file could not be found at:\n" + helpFilePath, "File Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var helpContent = File.ReadAllText(helpFilePath);
                var topics = System.Text.Json.JsonSerializer.Deserialize<List<HelpTopic>>(helpContent);

                if (topics == null || topics.Count == 0)
                {
                    MessageBox.Show("No help topics found in the file.", "Help Empty", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Window helpWindow = new Window
                {
                    Title = "ADA Help & Documentation",
                    Width = 900,
                    Height = 650,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F0F0F0"))
                };

                Grid grid = new Grid { Margin = new Thickness(10) };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                ListBox listBox = new ListBox
                {
                    ItemsSource = topics,
                    DisplayMemberPath = "Title",
                    BorderThickness = new Thickness(1),
                    Background = System.Windows.Media.Brushes.White,
                    FontSize = 14,
                    Padding = new Thickness(5)
                };
                
                ScrollViewer scroll = new ScrollViewer { Margin = new Thickness(10, 0, 0, 0) };
                TextBlock textBlock = new TextBlock
                {
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 16,
                    Padding = new Thickness(15),
                    Background = System.Windows.Media.Brushes.White,
                    LineHeight = 24
                };
                scroll.Content = textBlock;

                listBox.SelectionChanged += (s, ev) =>
                {
                    if (listBox.SelectedItem is HelpTopic selectedTopic)
                    {
                        textBlock.Text = selectedTopic.Content;
                    }
                };

                Grid.SetColumn(listBox, 0);
                Grid.SetColumn(scroll, 1);
                grid.Children.Add(listBox);
                grid.Children.Add(scroll);

                helpWindow.Content = grid;
                
                if (topics.Count > 0) listBox.SelectedIndex = 0;

                helpWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading help topics: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private string GetGithubToken()
        {
            string token = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
            if (!string.IsNullOrEmpty(token)) return token;

            string localFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "github_token.txt");
            if (File.Exists(localFile)) return File.ReadAllText(localFile).Trim();

            string userFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".afp_github_token");
            if (File.Exists(userFile)) return File.ReadAllText(userFile).Trim();

            return null;
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
                string githubToken = GetGithubToken();
                if (string.IsNullOrEmpty(githubToken))
                {
                    aiResponseTxt.Text = "No GitHub token found. Please do one of the following:\n" +
                                         "1. Set a 'GITHUB_TOKEN' environment variable.\n" +
                                         "2. Create a 'github_token.txt' file in the app directory.\n" +
                                         $"3. Create a '.afp_github_token' file in your user folder ({Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}).";
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

        private async void AskAIGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            var node = btn?.Tag as AfpNode;
            var aiResponseTxt = this.FindName("AiResponseTxt") as TextBox;
            
            if (node == null || aiResponseTxt == null || node.Children == null || node.Children.Count == 0) return;

            btn.IsEnabled = false;
            btn.Content = "✨ Thinking...";
            aiResponseTxt.Visibility = Visibility.Visible;
            aiResponseTxt.Text = "Generating group explanation...\n";

            try
            {
                string githubToken = GetGithubToken();
                if (string.IsNullOrEmpty(githubToken))
                {
                    aiResponseTxt.Text = "No GitHub token found. Please do one of the following:\n" +
                                         "1. Set a 'GITHUB_TOKEN' environment variable.\n" +
                                         "2. Create a 'github_token.txt' file in the app directory.\n" +
                                         $"3. Create a '.afp_github_token' file in your user folder ({Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}).";
                    return;
                }
                
                var client = new OpenAI.OpenAIClient(new ApiKeyCredential(githubToken), new OpenAI.OpenAIClientOptions { Endpoint = new Uri("https://models.inference.ai.azure.com") });
                var chatClient = client.GetChatClient("gpt-4o-mini");

                // Provide a high-level summary of the group's contents
                var summaryList = node.Children.Take(50).Select(c => $"- {c.Name}: {c.FriendlyName}");
                string childrenContext = string.Join("\n", summaryList);
                if (node.Children.Count > 50)
                {
                    childrenContext += $"\n... and {node.Children.Count - 50} more fields (truncated for context limits).";
                }

                string prompt = $@"
You are an expert in the IBM AFP (Advanced Function Presentation) Datastream specification.
I am analyzing an AFP group that starts with the '{node.Name}' ({node.DisplayName}) structured field. 
Here are the nested structured fields contained within this group:
{childrenContext}

In 3-4 concise bullet points, explain the primary purpose of this specific group type in an AFP datastream and summarize what this specific sequence of fields is accomplishing at a high level.";

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
                btn.Content = "✨ Explain Group Context";
            }
        }

        private async void SendAIBtn_Click(object sender, RoutedEventArgs e)
        {
            await ProcessChatInput();
        }

        private async void ChatInputBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) != System.Windows.Input.ModifierKeys.Shift)
            {
                e.Handled = true;
                await ProcessChatInput();
            }
        }

        private async Task ProcessChatInput()
        {
            var chatInputBox = this.FindName("ChatInputBox") as TextBox;
            var chatScrollViewer = this.FindName("ChatScrollViewer") as ScrollViewer;
            var sendAIBtn = this.FindName("SendAIBtn") as Button;

            if (chatInputBox == null || string.IsNullOrWhiteSpace(chatInputBox.Text) || sendAIBtn == null) return;

            string userText = chatInputBox.Text.Trim();
            chatInputBox.Text = "";
            chatInputBox.IsEnabled = false;
            sendAIBtn.IsEnabled = false;

            // Add user message to UI
            _chatHistoryUI.Add(new ChatMessageModel
            {
                Text = userText,
                BackgroundBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E1F5FE")),
                ForegroundBrush = System.Windows.Media.Brushes.Black,
                Alignment = HorizontalAlignment.Right
            });

            _llmChatHistory.Add(new UserChatMessage(userText));

            // Scroll to bottom
            chatScrollViewer?.ScrollToBottom();

            // Create loading placeholder
            var aiResponseUi = new ChatMessageModel
            {
                Text = "...",
                BackgroundBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F8FF")),
                ForegroundBrush = System.Windows.Media.Brushes.Black,
                Alignment = HorizontalAlignment.Left
            };
            _chatHistoryUI.Add(aiResponseUi);
            chatScrollViewer?.ScrollToBottom();

            try
            {
                string githubToken = GetGithubToken();
                if (string.IsNullOrEmpty(githubToken))
                {
                    aiResponseUi.Text = "No GitHub token found. Please securely configure your token first.";
                    return;
                }

                var client = new OpenAI.OpenAIClient(new ApiKeyCredential(githubToken), new OpenAI.OpenAIClientOptions { Endpoint = new Uri("https://models.inference.ai.azure.com") });
                var chatClient = client.GetChatClient("gpt-4o-mini");

                var responseStream = chatClient.CompleteChatStreamingAsync(_llmChatHistory);

                aiResponseUi.Text = "";
                string fullResponse = "";
                await foreach (var update in responseStream)
                {
                    if (update.ContentUpdate != null)
                    {
                        foreach (var part in update.ContentUpdate)
                        {
                            fullResponse += part.Text;
                            aiResponseUi.Text = fullResponse;
                            chatScrollViewer?.ScrollToBottom();
                        }
                    }
                }
                
                _llmChatHistory.Add(new AssistantChatMessage(fullResponse));
            }
            catch (Exception ex)
            {
                aiResponseUi.Text = $"Error getting AI response: {ex.Message}";
                aiResponseUi.ForegroundBrush = System.Windows.Media.Brushes.DarkRed;
            }
            finally
            {
                chatInputBox.IsEnabled = true;
                sendAIBtn.IsEnabled = true;
                chatInputBox.Focus();
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

            // --- Font Mapping (Path A) ---
            var fontMap = new Dictionary<byte, (string Name, double Size, FontWeight Weight, FontStyle Style)>();

            string MapFocaFont(string focaName, out double sizeDips)
            {
                double ptSize = 10.0; // default point size
                string fn = focaName.ToUpper().Trim();
                
                // Common IBM FOCA naming conventions often embed pitch or point size
                if (fn.Contains("10")) ptSize = 12.0; // 10 pitch typically maps to ~12pt
                else if (fn.Contains("12")) ptSize = 10.0; // 12 pitch maps to ~10pt
                else if (fn.Contains("15")) ptSize = 8.0;
                else if (fn.Contains("20")) ptSize = 6.0;

                string mappedSystemFont = SettingsService.CurrentSettings.DefaultFont;

                foreach (var mapping in SettingsService.CurrentSettings.FontMappings)
                {
                    if (fn.Contains(mapping.MatchString.ToUpper()) || fn.StartsWith(mapping.MatchString.ToUpper()))
                    {
                        mappedSystemFont = mapping.SystemFontName;
                        break;
                    }
                }

                // If the user's AFP has specific sizes in the name (like X0something07), try to catch them here if they aren't caught by the 0x1F triplet
                if (fn.Contains("06")) ptSize = 6.0;
                else if (fn.Contains("07")) ptSize = 7.0;
                else if (fn.Contains("08")) ptSize = 8.0;
                else if (fn.Contains("09")) ptSize = 9.0;

                // WPF FontSize expects Device Independent Pixels (1/96 inch). 
                // Typographic points are 1/72 inch. Therefore, to display 10pt text in WPF, FontSize must be 10 * (96/72).
                sizeDips = ptSize * (96.0 / 72.0);
                
                return mappedSystemFont;
            }

            void BuildFontMap(AfpNode node)
            {
                if (node.Id == "D3AB8A") // MCF - Map Coded Font
                {
                    try
                    {
                        using (var reader = new BinaryReader(File.OpenRead(_currentFilePath)))
                        {
                            reader.BaseStream.Seek(node.Offset + 9, SeekOrigin.Begin);
                            byte[] mcfData = reader.ReadBytes(node.Length);
                            
                            int pos = 0;
                            while (pos + 1 < mcfData.Length) // Needs at least 2 bytes for rgLen
                            {
                                int rgLen = (mcfData[pos] << 8) | mcfData[pos + 1];
                                if (rgLen < 2 || pos + rgLen > mcfData.Length) break;

                                byte localId = 0;
                                string fontName = "Courier New";
                                double fontSize = 10.0;
                                FontWeight fontWeight = FontWeights.Normal;
                                FontStyle fontStyle = FontStyles.Normal;

                                // Scan triplets within this mapped group
                                int tripPos = pos + 2;
                                int groupEnd = pos + rgLen;
                                while (tripPos + 1 < groupEnd)
                                {
                                    int tLen = mcfData[tripPos];
                                    if (tLen < 2 || tripPos + tLen > groupEnd) break;
                                    byte tId = mcfData[tripPos + 1];

                                    if (tId == 0x24 && tLen >= 4) // Resource Local Identifier
                                    {
                                        localId = mcfData[tripPos + 3];
                                    }
                                    else if (tId == 0x02 && tLen >= 4) // Fully Qualified Name (FQN)
                                    {
                                        byte fqnType = mcfData[tripPos + 2];
                                        // 0x8E = Coded Font, 0x86 = Font Character Set, 0x85 = Code Page
                                        if (fqnType == 0x8E || fqnType == 0x86)
                                        {
                                            byte fqnFmt = mcfData[tripPos + 3]; // Usually 0x00 for character string
                                            if (fqnFmt == 0x00 && tLen > 4)
                                            {
                                                string focaName = ebcdic.GetString(mcfData, tripPos + 4, tLen - 4).Trim();
                                                fontName = MapFocaFont(focaName, out fontSize);
                                            }
                                        }
                                    }
                                    else if (tId == 0x1F && tLen >= 20) // Font Descriptor Specification
                                    {
                                        byte ftWtClass = mcfData[tripPos + 2];
                                        fontWeight = ftWtClass switch
                                        {
                                            0x01 => FontWeights.UltraLight,
                                            0x02 => FontWeights.ExtraLight,
                                            0x03 => FontWeights.Light,
                                            0x04 => FontWeights.SemiBold, // Or Medium? Let's stick with closest mapping
                                            0x05 => FontWeights.Normal,
                                            0x06 => FontWeights.SemiBold,
                                            0x07 => FontWeights.Bold,
                                            0x08 => FontWeights.ExtraBold,
                                            0x09 => FontWeights.UltraBold,
                                            _ => FontWeights.Normal
                                        };

                                        int ftHeight = (mcfData[tripPos + 4] << 8) | mcfData[tripPos + 5];
                                        if (ftHeight > 0)
                                        {
                                            // The value is in 1440ths of an inch.
                                            // 1 point = 1/72 inch = 20/1440.
                                            // Therefore, typographical point size is ftHeight / 20.0
                                            double pts = ftHeight / 20.0;
                                            
                                            // Convert those typographical points (1/72) to WPF Device Independent Pixels (1/96)
                                            fontSize = pts * (96.0 / 72.0);
                                        }

                                        byte ftDsFlags = mcfData[tripPos + 8];
                                        if ((ftDsFlags & 0x80) == 0x80) // Bit 0 is Italic
                                        {
                                            fontStyle = FontStyles.Italic;
                                        }
                                    }
                                    tripPos += tLen;
                                }

                                fontMap[localId] = (fontName, fontSize, fontWeight, fontStyle);
                                pos += rgLen;
                            }
                        }
                    }
                    catch { }
                }
                foreach (var child in node.Children) BuildFontMap(child);
            }

            // Extract fonts from the entire document context (or just the active environment groups)
            foreach (var docRoot in _flatNodeList.Where(n => n.Parent == null))
            {
                BuildFontMap(docRoot);
            }
            // -----------------------------

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

            string currentFontName = "Courier New";
            double currentFontSize = 10.0;
            FontWeight currentFontWeight = FontWeights.Normal;
            FontStyle currentFontStyle = FontStyles.Normal;
            SolidColorBrush currentTextColor = System.Windows.Media.Brushes.Black;

            SolidColorBrush GetOcaSolidColor(int ocaVal)
            {
                return ocaVal switch
                {
                    1 => System.Windows.Media.Brushes.Blue,
                    2 => System.Windows.Media.Brushes.Red,
                    3 => System.Windows.Media.Brushes.Magenta,
                    4 => System.Windows.Media.Brushes.Green,
                    5 => System.Windows.Media.Brushes.Cyan,
                    6 => System.Windows.Media.Brushes.Yellow,
                    7 => System.Windows.Media.Brushes.White,
                    8 => System.Windows.Media.Brushes.Black,
                    16 => System.Windows.Media.Brushes.Brown,
                    _ => System.Windows.Media.Brushes.Black
                };
            }

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
                                    short val = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                    currentY = (val / textResY) * 96.0;
                                }
                                else if (csCode == 0xC7 || csCode == 0xC6) // AMI
                                {
                                    short val = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
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
                                else if (csCode == 0x75 || csCode == 0x74) // STC (Set Text Color)
                                {
                                    if (dataLen >= 2)
                                    {
                                        int colorVal = (payload[dataStart] << 8) | payload[dataStart + 1];
                                        currentTextColor = GetOcaSolidColor(colorVal);
                                    }
                                }
                                else if (csCode == 0x81 || csCode == 0x80) // SEC (Set Extended Text Color)
                                {
                                    if (dataLen >= 12)
                                    {
                                        byte colSpace = payload[dataStart + 1];
                                        if (colSpace == 0x01) // RGB
                                        {
                                            byte rDim = payload[dataStart + 6];
                                            byte gDim = payload[dataStart + 7];
                                            byte bDim = payload[dataStart + 8];
                                            if (rDim == 8 && gDim == 8 && bDim == 8 && dataLen >= 13)
                                            {
                                                byte r = payload[dataStart + 10];
                                                byte g = payload[dataStart + 11];
                                                byte b = payload[dataStart + 12];
                                                currentTextColor = new SolidColorBrush(Color.FromRgb(r, g, b));
                                            }
                                        }
                                        else if (colSpace == 0x04) // CMYK
                                        {
                                            byte cDim = payload[dataStart + 6];
                                            byte mDim = payload[dataStart + 7];
                                            byte yDim = payload[dataStart + 8];
                                            byte kDim = payload[dataStart + 9];
                                            if (cDim == 8 && mDim == 8 && yDim == 8 && kDim == 8 && dataLen >= 14)
                                            {
                                                byte c = payload[dataStart + 10];
                                                byte m = payload[dataStart + 11];
                                                byte y = payload[dataStart + 12];
                                                byte k = payload[dataStart + 13];
                                                byte r = (byte)(255 * (1 - c / 255.0) * (1 - k / 255.0));
                                                byte gg = (byte)(255 * (1 - m / 255.0) * (1 - k / 255.0));
                                                byte bb = (byte)(255 * (1 - y / 255.0) * (1 - k / 255.0));
                                                currentTextColor = new SolidColorBrush(Color.FromRgb(r, gg, bb));
                                            }
                                        }
                                        else if (colSpace == 0x08) // Standard OCA
                                        {
                                            if (dataLen >= 13)
                                            {
                                                int colorVal = (payload[dataStart + 10] << 8) | payload[dataStart + 11];
                                                currentTextColor = GetOcaSolidColor(colorVal);
                                            }
                                        }
                                    }
                                }
                                else if (csCode == 0xF1 || csCode == 0xF0) // SCFL (Set Coded Font Local)
                                {
                                    if (dataLen >= 1)
                                    {
                                        byte localId = payload[dataStart];
                                        if (fontMap.TryGetValue(localId, out var mappedFont))
                                        {
                                            currentFontName = mappedFont.Name;
                                            currentFontSize = mappedFont.Size;
                                            currentFontWeight = mappedFont.Weight;
                                            currentFontStyle = mappedFont.Style;
                                        }
                                    }
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
                                            FontFamily = new FontFamily(currentFontName),
                                            FontSize = currentFontSize,
                                            FontWeight = currentFontWeight,
                                            FontStyle = currentFontStyle,
                                            Foreground = currentTextColor
                                        };
                                        Canvas.SetLeft(tb, currentX);
                                        // A slight offset adjustment since 0,0 for text is usually top-left in WPF but baseline in AFP
                                        Canvas.SetTop(tb, currentY - currentFontSize);
                                        docCanvas.Children.Add(tb);
                                        textBlocksAdded++;
                                        
                                        // Auto-increment currentX by the actual measured width of the text block we just drew.
                                        // This prevents overlap when consecutive TRN sequences or inline state changes occur.
                                        tb.Measure(new Size(Double.PositiveInfinity, Double.PositiveInfinity));
                                        currentX += tb.DesiredSize.Width; 
                                    }
                                }
                                else if (csCode == 0xE7 || csCode == 0xE6) // DBR (Draw B-axis Rule)
                                {
                                    if (dataLen >= 2)
                                    {
                                        short ruleLength = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                        double rLength = (ruleLength / textResY) * 96.0;
                                        if (ruleLength == 0x7FFF) rLength = (canvasHeight > currentY) ? (canvasHeight - currentY) : 0; // special "draw to margin" value

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
                                            if (wVal1 == 0x7FFF) rWidth = (canvasWidth > currentX) ? (canvasWidth - currentX) : 0;
                                        }

                                        System.Windows.Shapes.Rectangle rect = new System.Windows.Shapes.Rectangle
                                        {
                                            Width = Math.Max(1.0, Math.Abs(rWidth)),
                                            Height = Math.Max(1.0, Math.Abs(rLength)),
                                            Fill = currentTextColor
                                        };
                                        Canvas.SetLeft(rect, rWidth < 0 ? currentX + rWidth : currentX);
                                        Canvas.SetTop(rect, rLength < 0 ? currentY + rLength : currentY);
                                        // Ensure rules rendered early don't completely trap text if there is overlap
                                        Panel.SetZIndex(rect, -1);
                                        docCanvas.Children.Add(rect);
                                    }
                                }
                                else if (csCode == 0xE5 || csCode == 0xE4) // DIR (Draw I-axis Rule)
                                {
                                    if (dataLen >= 2)
                                    {
                                        short ruleLength = (short)((payload[dataStart] << 8) | payload[dataStart + 1]);
                                        double rLength = (ruleLength / textResX) * 96.0;
                                        if (ruleLength == 0x7FFF) rLength = (canvasWidth > currentX) ? (canvasWidth - currentX) : 0; // special "draw to margin" value

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
                                            if (wVal1 == 0x7FFF) rWidth = (canvasHeight > currentY) ? (canvasHeight - currentY) : 0;
                                        }

                                        System.Windows.Shapes.Rectangle rect = new System.Windows.Shapes.Rectangle
                                        {
                                            Width = Math.Max(1.0, Math.Abs(rLength)),
                                            Height = Math.Max(1.0, Math.Abs(rWidth)),
                                            Fill = currentTextColor
                                        };
                                        Canvas.SetLeft(rect, rLength < 0 ? currentX + rLength : currentX);
                                        Canvas.SetTop(rect, rWidth < 0 ? currentY + rWidth : currentY);
                                        // Ensure rules rendered early don't completely trap text if there is overlap
                                        Panel.SetZIndex(rect, -1);
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

                                    // First check IOB triplets for an Object Area Size override (0x4C)
                                    // IOB payload typically has fixed params for the first 38 bytes or so, but let's scan for triplets based on spec.
                                    // IOB triplets start at offset 24 (or 26 in some older versions? Wait, standard is:
                                    // Obj Name: 0-7, Reserved: 8, Obj Type: 9, XoaOset: 10-12, YoaOset: 13-15, Rot: 16-17..
                                    // Let's just find the 0x4C safely since it's specific to AFP. A full parser would use the offset provided in byte 24 or similar, 
                                    // but we can just use the PGD/Active Environment Group units if found, or default to the resX/resY
                                    bool foundIOB4C = false;
                                    // The IOB triplet length starts at offset 26.
                                    if (payload.Length > 26)
                                    {
                                        int iobTripStart = 26; // Default 
                                        int iidx = iobTripStart;
                                        double iobUnitsX = resX, iobUnitsY = resY; 
                                        double iobXMult = 10.0, iobYMult = 10.0;

                                        while (iidx + 1 < payload.Length)
                                        {
                                            int tLen = payload[iidx];
                                            if (tLen < 2 || iidx + tLen > payload.Length) break;
                                            byte tId = payload[iidx + 1];

                                            if (tId == 0x4B && tLen >= 8) // Measurement override on IOB
                                            {
                                                byte xBase = payload[iidx + 2];
                                                byte yBase = payload[iidx + 3];
                                                iobXMult = xBase == 0x01 ? 3.93701 : 10.0;
                                                iobYMult = yBase == 0x01 ? 3.93701 : 10.0;
                                                iobUnitsX = (payload[iidx + 4] << 8) | payload[iidx + 5];
                                                iobUnitsY = (payload[iidx + 6] << 8) | payload[iidx + 7];
                                            }
                                            else if (tId == 0x4C && tLen >= 9) // Area Size triplet on IOB
                                            {
                                                int xSize = (payload[iidx + 3] << 16) | (payload[iidx + 4] << 8) | payload[iidx + 5];
                                                int ySize = (payload[iidx + 6] << 16) | (payload[iidx + 7] << 8) | payload[iidx + 8];
                                                
                                                if (iobUnitsX > 0 && iobUnitsY > 0)
                                                {
                                                    targetWidth = (xSize * iobXMult / iobUnitsX) * 96.0;
                                                    targetHeight = (ySize * iobYMult / iobUnitsY) * 96.0;
                                                    foundIOB4C = true;
                                                }
                                            }
                                            iidx += tLen;
                                        }
                                    }

                                    // If not defined by IOB, fallback to Extract OBD for intrinsic scaling
                                    if (!foundIOB4C && targetBoc.Parent != null)
                                    {
                                        foreach (var sibling in targetBoc.Parent.Children)
                                        {
                                            if (sibling.Name == "OBD" && sibling.Offset < targetBoc.Offset)
                                            {
                                                obdNode = sibling;
                                            }
                                        }
                                    }

                                    if (!foundIOB4C && obdNode != null)
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
                                            int extractLen = Math.Min(9, bocFullPayload.Length - i);
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

                                                    // Use Stretch.Fill to ensure the image precisely hits the target boundaries specified by AFP
                                                    var img = new Image { Source = bitmap, Stretch = Stretch.Fill };
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

    public class HelpTopic
    {
        [System.Text.Json.Serialization.JsonPropertyName("Title")]
        public string Title { get; set; }
        
        [System.Text.Json.Serialization.JsonPropertyName("Content")]
        public string Content { get; set; }
    }
}
