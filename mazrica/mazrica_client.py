"""
Mazrica API クライアント
案件一覧・商品内訳情報を取得するためのクライアント
"""
import time
import requests
from typing import Optional, Generator
from dataclasses import dataclass

from .config import Config


@dataclass
class ProductDetail:
    """商品内訳情報"""
    product_id: Optional[int]
    product_name: str
    quantity: Optional[float]
    unit_price: Optional[float]
    amount: Optional[float]
    custom_fields: dict


@dataclass
class Deal:
    """案件情報"""
    id: int
    name: str
    customer_name: Optional[str]
    customer_id: Optional[int]
    deal_type_id: Optional[int]
    deal_type_name: Optional[str]
    phase_name: Optional[str]
    amount: Optional[float]
    expected_contract_date: Optional[str]
    created_at: str
    updated_at: str
    user_name: Optional[str]
    product_details: list[ProductDetail]
    custom_fields: dict


class MazricaAPIError(Exception):
    """Mazrica API エラー"""
    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        self.message = message
        super().__init__(f"Mazrica API Error ({status_code}): {message}")


class MazricaClient:
    """Mazrica APIクライアント"""
    
    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None):
        self.api_key = api_key or Config.MAZRICA_API_KEY
        self.base_url = base_url or Config.MAZRICA_BASE_URL
        self.rate_limit = Config.API_RATE_LIMIT
        self._last_request_time = 0.0
        
        if not self.api_key:
            raise ValueError("Mazrica API Key is required")
    
    def _wait_for_rate_limit(self):
        """レート制限を守るための待機"""
        elapsed = time.time() - self._last_request_time
        if elapsed < self.rate_limit:
            time.sleep(self.rate_limit - elapsed)
        self._last_request_time = time.time()
    
    def _request(self, method: str, endpoint: str, params: Optional[dict] = None) -> dict:
        """APIリクエストを実行"""
        self._wait_for_rate_limit()
        
        url = f"{self.base_url}{endpoint}"
        headers = {
            "X-Api-Key": self.api_key,
            "Content-Type": "application/json"
        }
        
        response = requests.request(
            method=method,
            url=url,
            headers=headers,
            params=params
        )
        
        if response.status_code == 429:
            # レート制限に達した場合は待機してリトライ
            time.sleep(1)
            return self._request(method, endpoint, params)
        
        if response.status_code >= 400:
            raise MazricaAPIError(
                status_code=response.status_code,
                message=response.text
            )
        
        return response.json()
    
    def get_deal_types(self) -> list[dict]:
        """案件タイプ一覧を取得"""
        result = self._request("GET", "/deal_types")
        return result.get("dealTypes", [])
    
    def get_deals(
        self,
        deal_type_id: Optional[int] = None,
        page: int = 1,
        limit: int = 100,
        sort: str = "-updatedAt"
    ) -> dict:
        """
        案件一覧を取得
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
            page: ページ番号
            limit: 1ページあたりの件数
            sort: ソート順（デフォルト: 更新日時降順）
        
        Returns:
            {"deals": [...], "totalCount": int, "page": int}
        """
        params = {
            "page": page,
            "limit": limit,
            "sort": sort
        }
        
        if deal_type_id:
            params["dealTypeId"] = deal_type_id
        
        return self._request("GET", "/deals", params)
    
    def get_all_deals(
        self,
        deal_type_id: Optional[int] = None,
        limit_per_page: int = 100
    ) -> Generator[dict, None, None]:
        """
        全案件を取得（ページネーション対応）
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
            limit_per_page: 1ページあたりの件数
        
        Yields:
            案件データ
        """
        page = 1
        while True:
            result = self.get_deals(
                deal_type_id=deal_type_id,
                page=page,
                limit=limit_per_page
            )
            
            deals = result.get("deals", [])
            if not deals:
                break
            
            for deal in deals:
                yield deal
            
            total_count = result.get("totalCount", 0)
            if page * limit_per_page >= total_count:
                break
            
            page += 1
    
    def _parse_product_detail(self, data: dict) -> ProductDetail:
        """商品内訳データをパース"""
        return ProductDetail(
            product_id=data.get("productId"),
            product_name=data.get("productName", ""),
            quantity=data.get("quantity"),
            unit_price=data.get("unitPrice"),
            amount=data.get("amount"),
            custom_fields=data.get("customFields", {})
        )
    
    def _parse_deal(self, data: dict) -> Deal:
        """案件データをパース"""
        # 取引先情報
        customer = data.get("customer", {}) or {}
        
        # 案件タイプ情報
        deal_type = data.get("dealType", {}) or {}
        
        # フェーズ情報
        phase = data.get("phase", {}) or {}
        
        # 担当者情報
        user = data.get("user", {}) or {}
        
        # 商品内訳
        product_details = [
            self._parse_product_detail(pd)
            for pd in data.get("productDetails", [])
        ]
        
        return Deal(
            id=data.get("id"),
            name=data.get("name", ""),
            customer_name=customer.get("name"),
            customer_id=customer.get("id"),
            deal_type_id=deal_type.get("id"),
            deal_type_name=deal_type.get("name"),
            phase_name=phase.get("name"),
            amount=data.get("amount"),
            expected_contract_date=data.get("expectedContractDate"),
            created_at=data.get("createdAt", ""),
            updated_at=data.get("updatedAt", ""),
            user_name=user.get("name"),
            product_details=product_details,
            custom_fields=data.get("customFields", {})
        )
    
    def fetch_deals_with_products(
        self,
        deal_type_id: Optional[int] = None
    ) -> list[Deal]:
        """
        案件一覧を商品内訳付きで取得
        
        Args:
            deal_type_id: 案件タイプIDでフィルタ
        
        Returns:
            案件リスト
        """
        deals = []
        for deal_data in self.get_all_deals(deal_type_id=deal_type_id):
            deal = self._parse_deal(deal_data)
            deals.append(deal)
        
        return deals


# テスト用
if __name__ == "__main__":
    from .config import load_dotenv
    load_dotenv()
    
    client = MazricaClient()
    
    # 案件タイプ一覧を取得
    print("=== 案件タイプ一覧 ===")
    deal_types = client.get_deal_types()
    for dt in deal_types:
        print(f"  ID: {dt.get('id')}, Name: {dt.get('name')}")
    
    # 案件一覧を取得（最初の10件）
    print("\n=== 案件一覧（最初の10件） ===")
    result = client.get_deals(limit=10)
    for deal_data in result.get("deals", [])[:10]:
        deal = client._parse_deal(deal_data)
        print(f"  ID: {deal.id}, Name: {deal.name}, Customer: {deal.customer_name}")
        if deal.product_details:
            for pd in deal.product_details:
                print(f"    - Product: {pd.product_name}, Amount: {pd.amount}")

