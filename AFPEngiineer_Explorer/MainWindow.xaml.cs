using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace AFPEngineer_Explorer
{
    public partial class MainWindow : Window
    {
        // EBCDIC (Code Page 37) is the standard for most AFP files
        private static Encoding ebcdic = Encoding.GetEncoding("IBM037");
        public MainWindow()
        {
            // Add this line to unlock EBCDIC support
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

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

        private void ParseAfpFile(string filePath)
        {
            // Bypass the generator: Find the controls manually
            var tree = this.FindName("AfpTreeView") as TreeView;
            var canvas = this.FindName("PageCanvas") as Canvas;

            if (tree == null) return;

            tree.Items.Clear();
            Stack<TreeViewItem> containerStack = new Stack<TreeViewItem>();

            try
            {
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

                            byte[] dataPayload = reader.ReadBytes(length - 5);
                            // Instead of Encoding.ASCII, we try to decode as EBCDIC
                            string textPreview = ebcdic.GetString(dataPayload);

                            // Clean up non-printable characters
                            textPreview = Regex.Replace(textPreview, @"[^\u0020-\u007E\u00A0-\u00FF]", ".");
                            TreeViewItem newItem = new TreeViewItem();
                            newItem.Header = $"{hexId} - {LookupAfpName(hexId)}";

                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine($"Field ID: {hexId}");
                            sb.AppendLine($"Offset: 0x{(reader.BaseStream.Position - length - 1):X}");
                            sb.AppendLine($"[EXPERT NOTE]: {GetExpertAdvice(hexId)}");
                            sb.AppendLine($"\n[DATA PEEK]: {textPreview}");
                            newItem.Tag = sb.ToString();

                            // Resize Canvas if PGD is found
                            if (hexId == "D3A6AF" && dataPayload.Length >= 15 && canvas != null)
                            {
                                int width = (dataPayload[9] << 16) | (dataPayload[10] << 8) | dataPayload[11];
                                int height = (dataPayload[12] << 16) | (dataPayload[13] << 8) | dataPayload[14];
                                canvas.Width = width / 15.0;
                                canvas.Height = height / 15.0;
                                canvas.Children.Clear();
                            }

                            // Nesting Logic
                            if (hexId.StartsWith("D3A8"))
                            {
                                if (containerStack.Count > 0) containerStack.Peek().Items.Add(newItem);
                                else tree.Items.Add(newItem);
                                containerStack.Push(newItem);
                                newItem.IsExpanded = true;
                            }
                            else if (hexId.StartsWith("D3A9"))
                            {
                                if (containerStack.Count > 0) containerStack.Pop();
                            }
                            else
                            {
                                if (containerStack.Count > 0) containerStack.Peek().Items.Add(newItem);
                                else tree.Items.Add(newItem);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private string LookupAfpName(string id) => id switch
        {
            "D3A8A8" => "Begin Document (BDT)",
            "D3A9A8" => "End Document (EDT)",
            "D3A8AF" => "Begin Page (BPG)",
            "D3A9AF" => "End Page (EPG)",
            "D3A6AF" => "Page Descriptor (PGD)",
            "D3EEBB" => "Presentation Text (PTX)",
            _ => "Other Structured Field"
        };

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
        private void AfpTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var fieldNameTxt = this.FindName("FieldNameTxt") as TextBlock;
            var hexDataTxt = this.FindName("HexDataTxt") as TextBlock;
            var tree = sender as TreeView;

            if (tree?.SelectedItem is TreeViewItem selected)
            {
                if (fieldNameTxt != null) fieldNameTxt.Text = selected.Header.ToString();
                if (hexDataTxt != null) hexDataTxt.Text = selected.Tag.ToString();
            }
        }
    }
}