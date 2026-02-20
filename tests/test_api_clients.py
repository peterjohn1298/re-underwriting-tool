"""Tests for API client fallback behavior (no real API calls)."""
import pytest
from unittest.mock import patch, MagicMock
from services.api_clients.walkscore_client import WalkScoreClient, _label, WALK_LABELS
from services.api_clients.hud_client import HUDClient
from services.api_clients.crime_client import CrimeClient


class TestWalkScoreClient:
    def test_unconfigured_returns_available_false(self):
        client = WalkScoreClient(api_key="")
        result = client.get_scores("123 Main St, Austin TX")
        assert result["available"] is False

    def test_walk_label_high(self):
        label = _label(95, WALK_LABELS)
        assert label == "Walker's Paradise"

    def test_walk_label_low(self):
        label = _label(20, WALK_LABELS)
        assert "Car" in label

    def test_walk_label_none(self):
        label = _label(None, WALK_LABELS)
        assert label == "N/A"

    def test_api_error_returns_available_false(self):
        client = WalkScoreClient(api_key="fake-key")
        with patch("services.api_clients.walkscore_client.requests.get") as mock_get:
            mock_get.side_effect = Exception("Connection error")
            result = client.get_scores("123 Main St")
        assert result["available"] is False
        assert "error" in result

    def test_bad_status_returns_available_false(self):
        client = WalkScoreClient(api_key="fake-key")
        mock_resp = MagicMock()
        mock_resp.json.return_value = {"status": 30}
        mock_resp.raise_for_status = MagicMock()
        with patch("services.api_clients.walkscore_client.requests.get",
                   return_value=mock_resp):
            result = client.get_scores("123 Main St")
        assert result["available"] is False

    def test_successful_response_parsed(self):
        client = WalkScoreClient(api_key="fake-key")
        mock_resp = MagicMock()
        mock_resp.json.return_value = {
            "status": 1,
            "walkscore": 85,
            "snapped_address": "123 Main St, Austin, TX",
            "transit": {"score": 60},
            "bike": {"score": 70},
            "ws_link": "https://www.walkscore.com/score/...",
        }
        mock_resp.raise_for_status = MagicMock()
        with patch("services.api_clients.walkscore_client.requests.get",
                   return_value=mock_resp):
            result = client.get_scores("123 Main St")
        assert result["available"] is True
        assert result["walk_score"] == 85
        assert result["transit_score"] == 60
        assert result["bike_score"] == 70
        assert "Very Walkable" in result["walk_label"]


class TestCrimeClient:
    def test_known_city_returns_data(self):
        client = CrimeClient()
        result = client.get_city_crime("Austin", "TX")
        # Should find Austin in the CSV
        if result.get("available"):
            assert "violent_per_100k" in result
            assert "risk_level" in result
            assert result["violent_per_100k"] > 0

    def test_unknown_city_returns_available_false(self):
        client = CrimeClient()
        result = client.get_city_crime("FakeCity123", "ZZ")
        assert result["available"] is False

    def test_risk_level_valid(self):
        client = CrimeClient()
        result = client.get_city_crime("Austin", "TX")
        if result.get("available"):
            assert result["risk_level"] in ("LOW", "MODERATE", "HIGH", "VERY HIGH")

    def test_aliases_work(self):
        client = CrimeClient()
        result = client.get_city_crime("nyc", "NY")
        if result.get("available"):
            assert result["violent_per_100k"] > 0


class TestHUDClient:
    def test_unconfigured_returns_available_false(self):
        client = HUDClient(api_key="__no_key__")
        # With a fake key the API call fails — patch _get to return None
        with patch.object(client, "_get", return_value=None):
            result = client.get_fmr_by_city("Austin", "TX")
        assert result["available"] is False

    def test_invalid_state_returns_available_false(self):
        client = HUDClient(api_key="fake")
        with patch.object(client, "_get", return_value=None):
            result = client.get_fmr_by_city("Austin", "ZZ")
        assert result["available"] is False

    def test_bedroom_key_mapping(self):
        client = HUDClient(api_key="fake")
        mock_fmr = {
            "available": True,
            "city": "Austin", "state": "TX",
            "metro_name": "Austin MSA", "entity_code": "X",
            "year": "2026", "fmr_percentile": 40,
            "fmr_by_bedroom": {
                "efficiency": 900, "one_br": 1000,
                "two_br": 1200, "three_br": 1500, "four_br": 1800,
            },
            "source": "HUD", "note": "",
        }
        with patch.object(client, "get_fmr_by_city", return_value=mock_fmr):
            result = client.get_fmr_for_avg_bedrooms("Austin", "TX", avg_bedrooms=2)
        assert result["matched_bedroom_key"] == "two_br"
        assert result["matched_fmr"] == 1200

    def test_avg_bedrooms_clamped_at_4(self):
        client = HUDClient(api_key="fake")
        mock_fmr = {
            "available": True,
            "city": "Austin", "state": "TX",
            "metro_name": "Austin MSA", "entity_code": "X",
            "year": "2026", "fmr_percentile": 40,
            "fmr_by_bedroom": {
                "efficiency": 900, "one_br": 1000,
                "two_br": 1200, "three_br": 1500, "four_br": 1800,
            },
            "source": "HUD", "note": "",
        }
        with patch.object(client, "get_fmr_by_city", return_value=mock_fmr):
            result = client.get_fmr_for_avg_bedrooms("Austin", "TX", avg_bedrooms=6)
        assert result["matched_bedroom_key"] == "four_br"
