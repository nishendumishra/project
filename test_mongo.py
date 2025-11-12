import unittest
from unittest.mock import MagicMock, patch
from pymongo.collection import Collection
from bson import ObjectId

from cstgenai_common_services.support.mongo import MongoDBManager, MongoRepository


class TestMongoDBManager(unittest.TestCase):
    """Unit tests for MongoDBManager class."""

    @patch("cstgenai_common_services.support.mongo.MongoClient")
    def test_init_creates_client_and_db(self, mock_client):
        """Test MongoDBManager initializes MongoClient and sets DB."""
        mock_client_instance = MagicMock()
        mock_client.return_value = mock_client_instance
        mock_client_instance.__getitem__.return_value = "mock_db"

        manager = MongoDBManager("mongodb://localhost:27017", "test_db")

        mock_client.assert_called_once_with("mongodb://localhost:27017")
        self.assertEqual(manager.db, "mock_db")
        self.assertEqual(manager.client, mock_client_instance)

    @patch("cstgenai_common_services.support.mongo.MongoClient")
    def test_get_collection_returns_collection(self, mock_client):
        """Test get_collection returns a collection object."""
        mock_db = MagicMock()
        mock_collection = MagicMock(spec=Collection)
        mock_db.__getitem__.return_value = mock_collection
        mock_client.return_value.__getitem__.return_value = mock_db

        manager = MongoDBManager("mongodb://localhost:27017", "test_db")
        result = manager.get_collection("users")

        self.assertEqual(result, mock_collection)
        mock_db.__getitem__.assert_called_once_with("users")

    @patch("cstgenai_common_services.support.mongo.MongoClient")
    def test_close_closes_client(self, mock_client):
        """Test close() properly closes the MongoDB client."""
        mock_client_instance = MagicMock()
        mock_client.return_value = mock_client_instance

        manager = MongoDBManager("mongodb://localhost:27017", "test_db")
        manager.close()

        mock_client_instance.close.assert_called_once()


class TestMongoRepository(unittest.TestCase):
    """Unit tests for MongoRepository class."""

    def setUp(self):
        """Prepare mocks for MongoRepository tests."""
        self.mock_manager = MagicMock(spec=MongoDBManager)
        self.mock_collection = MagicMock(spec=Collection)
        self.mock_manager.get_collection.return_value = self.mock_collection

        class MockModel:
            def __init__(self, id=None, name=None):
                self.id = id
                self.name = name

        self.MockModel = MockModel
        self.repo = MongoRepository(self.mock_manager, "test_collection", self.MockModel)

    def test_init_assigns_collection_and_model(self):
        """Test __init__ assigns correct collection and model class."""
        self.assertEqual(self.repo.collection, self.mock_collection)
        self.assertEqual(self.repo.model_class, self.MockModel)

    def test_convert_to_model_with_id_conversion(self):
        """Test _convert_to_model converts document _id to id."""
        doc = {"_id": ObjectId("655b7a9b0f1c000000000000"), "name": "Alice"}
        model_instance = self.repo._convert_to_model(doc)

        self.assertEqual(model_instance.name, "Alice")
        self.assertIsInstance(model_instance.id, str)

    def test_find_all_returns_models(self):
        """Test find_all returns list of model instances."""
        self.mock_collection.find.return_value = [{"_id": "1", "name": "John"}]

        results = self.repo.find_all()
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].name, "John")
        self.mock_collection.find.assert_called_once()

    def test_find_all_with_limit(self):
        """Test find_all applies limit when provided."""
        mock_cursor = MagicMock()
        self.mock_collection.find.return_value = mock_cursor
        mock_cursor.limit.return_value = [{"_id": "2", "name": "Jane"}]

        results = self.repo.find_all(limit=1)
        mock_cursor.limit.assert_called_once_with(1)
        self.assertEqual(results[0].name, "Jane")

    def test_find_by_filter(self):
        """Test find_by_filter applies query and limit."""
        mock_cursor = MagicMock()
        self.mock_collection.find.return_value = mock_cursor
        mock_cursor.limit.return_value = [{"_id": "3", "name": "Doe"}]

        results = self.repo.find_by_filter({"name": "Doe"}, limit=1)
        mock_cursor.limit.assert_called_once_with(1)
        self.assertEqual(results[0].name, "Doe")

    def test_find_one_returns_model(self):
        """Test find_one returns single model object."""
        self.mock_collection.find_one.return_value = {"_id": "4", "name": "Bob"}
        result = self.repo.find_one({"name": "Bob"})
        self.assertEqual(result.name, "Bob")

    def test_find_one_returns_none_when_not_found(self):
        """Test find_one returns None if no result found."""
        self.mock_collection.find_one.return_value = None
        result = self.repo.find_one({"name": "Ghost"})
        self.assertIsNone(result)

    def test_find_by_id_returns_model(self):
        """Test find_by_id returns correct model."""
        self.mock_collection.find_one.return_value = {"_id": ObjectId(), "name": "Eve"}
        result = self.repo.find_by_id("655b7a9b0f1c000000000000")
        self.assertEqual(result.name, "Eve")

    def test_find_by_id_returns_none_when_not_found(self):
        """Test find_by_id returns None if document not found."""
        self.mock_collection.find_one.return_value = None
        result = self.repo.find_by_id("000000000000000000000000")
        self.assertIsNone(result)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
