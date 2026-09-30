"""List prices (USD per 1M tokens) for the Foundry deployments, from the Azure Retail Prices API
(prices.azure.com, East US, checked 2026-09-25). Your bill can differ (discounts, credits, currency).

Keyed by deployment name. gpt-6 prices are the short-context rates; their long-context rates are higher
and the cache-write charge isn't reported by the API usage figures, so gpt-6 costs are a floor.
"""
PRICES = {  # deployment: (input, cached input, output). Calls logged earlier on gpt-4 / o3-mini / DeepSeek keep
            # the cost stored when they were made.
    "gpt-4o": (2.50, 1.25, 10.00),        # gpt-4o 2024-11-20 Global
    "gpt-4o-mini": (0.165, 0.083, 0.66),  # Data Zone (US)
    "gpt-5-nano": (0.05, 0.005, 0.40),
    "gpt-6-sol": (2.00, 0.20, 10.00),
    "gpt-6-luna": (0.10, 0.01, 0.50),
}
ESTIMATES = {"gpt-6-sol", "gpt-6-luna"}  # see module docstring


def cost(model: str, input_tokens: int, cached_tokens: int, output_tokens: int) -> float | None:
    p = PRICES.get(model)
    if not p:
        return None
    return ((input_tokens - cached_tokens) * p[0] + cached_tokens * p[1] + output_tokens * p[2]) / 1e6
