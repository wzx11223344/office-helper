"""Auto-generated tests for office-helper."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import main


class TestMain:
    """Tests for office-helper module."""

    def test_module_import(self):
        """Test that main module imports correctly."""
        assert main is not None
        assert hasattr(main, "smart_file_classifier")


    def test_smart_file_classifier_basic(self):
        """Test smart file classifier."""
        result = main.smart_file_classifier("report", ["resume", "invoice", "contract"])
        assert "results" in result or "categories" in result

    def test_smart_file_classifier_empty(self):
        """Test classifier with no training data."""
        result = main.smart_file_classifier("test", [])
        assert result is not None

    def test_pdf_structure_analyzer_exists(self):
        """Test that pdf_structure_analyzer function is callable."""
        assert callable(main.pdf_structure_analyzer)
        assert main.pdf_structure_analyzer.__doc__ is not None


    def test_document_similarity_compare(self):
        """Test document similarity comparison."""
        result = main.document_similarity_compare("自然语言处理是人工智能的重要分支", "自然语言处理技术在搜索应用中广泛使用")
        assert "similarity" in result or "jaccard" in result or "cosine" in result

    def test_document_similarity_identical(self):
        """Test similarity of identical documents."""
        text = "测试文档内容"
        result = main.document_similarity_compare(text, text)
        assert result is not None

    def test_document_similarity_different(self):
        """Test similarity of very different documents."""
        result = main.document_similarity_compare("自然语言处理", "Python编程语言")
        assert result is not None

    def test_batch_rename_with_pattern_exists(self):
        """Test that batch_rename_with_pattern function is callable."""
        assert callable(main.batch_rename_with_pattern)
        assert main.batch_rename_with_pattern.__doc__ is not None

    def test_excel_data_merger_exists(self):
        """Test that excel_data_merger function is callable."""
        assert callable(main.excel_data_merger)
        assert main.excel_data_merger.__doc__ is not None
