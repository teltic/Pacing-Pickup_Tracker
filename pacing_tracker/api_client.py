"""Thin wrapper around the PriceLabs Customer API.

Base URL (https://api.pricelabs.co/v1) and the X-API-Key auth header are
confirmed working against the live account (see README).
"""

import os
import time

import requests

from . import config


class PriceLabsAPIError(RuntimeError):
    pass


class PriceLabsClient:
    def __init__(self, api_key=None, base_url=None, session=None, max_retries=3, backoff_seconds=2):
        self.api_key = api_key or os.environ.get(config.PRICELABS_API_KEY_ENV_VAR)
        if not self.api_key:
            raise PriceLabsAPIError(
                f"No API key found. Looked for a {config.PRICELABS_API_KEY_ENV_VAR} line in a "
                f".env file in the current folder ({os.getcwd()}) or a matching environment "
                "variable, and found neither. Create a .env file right next to this script "
                f"containing exactly one line: {config.PRICELABS_API_KEY_ENV_VAR}=your-key-here"
            )
        self.base_url = (base_url or config.PRICELABS_API_BASE_URL).rstrip("/")
        self.session = session or requests.Session()
        self.max_retries = max_retries
        self.backoff_seconds = backoff_seconds

    def _headers(self):
        return {"X-API-Key": self.api_key}

    def _request(self, method, path, params=None, json_body=None):
        url = f"{self.base_url}/{path.lstrip('/')}"
        last_error = None
        for attempt in range(self.max_retries):
            try:
                resp = self.session.request(
                    method, url, headers=self._headers(), params=params, json=json_body, timeout=30
                )
                if resp.status_code == 429 or resp.status_code >= 500:
                    last_error = PriceLabsAPIError(f"{resp.status_code} from {url}: {resp.text[:500]}")
                    time.sleep(self.backoff_seconds * (attempt + 1))
                    continue
                if not resp.ok:
                    raise PriceLabsAPIError(f"{resp.status_code} from {url}: {resp.text[:500]}")
                if not resp.content:
                    return {}  # e.g. a 204 No Content from DELETE
                return resp.json()
            except requests.RequestException as exc:
                last_error = PriceLabsAPIError(f"Request to {url} failed: {exc}")
                time.sleep(self.backoff_seconds * (attempt + 1))
        raise last_error

    def _get(self, path, params=None):
        return self._request("GET", path, params=params)

    def _post(self, path, json_body=None):
        return self._request("POST", path, json_body=json_body)

    def _delete(self, path, json_body=None):
        return self._request("DELETE", path, json_body=json_body)

    def get_report_builder_templates(self):
        return self._get("report_builder/templates")

    def get_report_builder_data(self, template_id):
        return self._post("report_builder/data", json_body={"template_id": template_id})

    def poll_report_builder_data(self, request_id):
        return self._post("report_builder/poll", json_body={"request_id": request_id})

    def get_listing_date_overrides(self, listing_id, pms, start_date=None, end_date=None):
        params = {"pms": pms}
        if start_date:
            params["start_date"] = start_date
        if end_date:
            params["end_date"] = end_date
        return self._get(f"listings/{listing_id}/overrides", params=params)

    def update_listing_date_overrides(self, listing_id, pms, overrides):
        return self._post(f"listings/{listing_id}/overrides", json_body={"pms": pms, "overrides": overrides})

    def delete_listing_date_overrides(self, listing_id, pms, dates):
        """Fully removes the override for each date (reverts to standard
        algorithmic pricing) -- confirmed live that this wipes ALL of a
        date's fields (min_stay, min/max price, etc.), not just price. Only
        safe to call for a date whose existing override has nothing else
        worth keeping; see push.py's _has_other_fields.
        """
        overrides = [{"date": d} for d in dates]
        return self._delete(f"listings/{listing_id}/overrides", json_body={"pms": pms, "overrides": overrides})
