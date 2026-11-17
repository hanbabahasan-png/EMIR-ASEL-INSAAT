from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path: sys.path.insert(0,str(ROOT))
import unittest
from unittest.mock import patch
import ai_gateway as ag


class AIGatewayTests(unittest.TestCase):
    def setUp(self):
        self.config = {
            "AI_PROVIDER": "openai",
            "AI_FALLBACK_ENABLED": "1",
            "AI_FALLBACK_ORDER": "openai,anthropic,gemini",
            "OPENAI_API_KEY": "test-openai-secret",
            "OPENAI_MODEL": "test-openai-model",
            "ANTHROPIC_API_KEY": "test-claude-secret",
            "ANTHROPIC_MODEL": "test-claude-model",
            "GEMINI_API_KEY": "test-gemini-secret",
            "GEMINI_MODEL": "test-gemini-model",
        }

    def test_provider_registry_requires_both_key_and_model(self):
        self.assertEqual(ag.configured_providers(self.config), ["openai", "anthropic", "gemini"])
        incomplete = dict(self.config, GEMINI_MODEL="")
        self.assertEqual(ag.configured_providers(incomplete), ["openai", "anthropic"])

    def test_status_never_exposes_api_secret(self):
        rows = ag.provider_status(self.config)
        self.assertEqual(len(rows), 4)
        rendered = repr(rows)
        self.assertNotIn("test-openai-secret", rendered)
        self.assertIn("Hazır", rendered)

    @patch("ai_gateway.call_provider")
    def test_auto_fallback_on_primary_provider_failure(self, call):
        call.side_effect = [RuntimeError("primary down"), ag.AIResult("ok", "anthropic", "test-claude-model")]
        result = ag.AIRouterClient(self.config).create(instructions="x", input="y")
        self.assertEqual(result.output_text, "ok")
        self.assertEqual(result.provider, "anthropic")
        self.assertTrue(result.fallback_used)
        self.assertEqual(call.call_count, 2)

    @patch("ai_gateway.call_provider")
    def test_fallback_can_be_disabled(self, call):
        cfg = dict(self.config, AI_FALLBACK_ENABLED="0")
        call.return_value = ag.AIResult("ok", "openai", "test-openai-model")
        ag.AIRouterClient(cfg).create(instructions="x", input="y")
        self.assertEqual(call.call_count, 1)

    def test_anthropic_and_gemini_multimodal_payload_translation(self):
        png_data_uri = "data:image/png;base64,aGVsbG8="
        messages = [{"role": "user", "content": [
            {"type": "input_text", "text": "Planı oku"},
            {"type": "input_image", "image_url": png_data_uri},
        ]}]
        system, anthropic_messages = ag._to_anthropic(ag._input_messages(messages))
        self.assertEqual(system, "")
        image_block = anthropic_messages[0]["content"][1]
        self.assertEqual(image_block["type"], "image")
        self.assertEqual(image_block["source"]["media_type"], "image/png")
        gemini_contents = ag._to_gemini_interactions(ag._input_messages(messages))
        self.assertEqual(gemini_contents[1]["type"], "image")
        self.assertEqual(gemini_contents[1]["mime_type"], "image/png")

    def test_unsupported_external_image_is_rejected_for_non_openai_provider(self):
        messages = [{"role": "user", "content": [{"type": "input_image", "image_url": "https://example.com/plan.png"}]}]
        with self.assertRaises(ValueError):
            ag._to_anthropic(ag._input_messages(messages))

    @patch("ai_gateway.call_provider")
    def test_multi_provider_review_keeps_partial_results(self, call):
        call.side_effect = [ag.AIResult('{"ok":true}', "openai", "m1"), RuntimeError("claude unavailable"), ag.AIResult("ok", "gemini", "m3")]
        results = ag.multi_provider_review(self.config, "system", "prompt", max_providers=3)
        self.assertEqual([x["status"] for x in results], ["success", "failed", "success"])
        self.assertEqual(results[2]["provider"], "gemini")


    def test_compatible_provider_requires_base_url(self):
        cfg = {"AI_COMPAT_API_KEY": "compat-secret", "AI_COMPAT_MODEL": "compat-model"}
        self.assertNotIn("openai_compatible", ag.configured_providers(cfg))
        cfg["AI_COMPAT_BASE_URL"] = "https://api.example/v1"
        self.assertIn("openai_compatible", ag.configured_providers(cfg))

    @patch("ai_gateway.requests.post")
    def test_gemini_uses_current_interactions_api_and_header_key(self, post):
        class Response:
            status_code = 200
            def raise_for_status(self): return None
            def json(self): return {"output_text": "Gemini test response"}
        post.return_value = Response()
        cfg = {"GEMINI_API_KEY": "gemini-private-key", "AI_TIMEOUT_SECONDS": "bad", "AI_MAX_OUTPUT_TOKENS": "bad"}
        text = ag._call_gemini(cfg, "gemini-test-model", "system", "hello")
        self.assertEqual(text, "Gemini test response")
        args, kwargs = post.call_args
        self.assertEqual(args[0], "https://generativelanguage.googleapis.com/v1/interactions")
        self.assertNotIn("gemini-private-key", args[0])
        self.assertEqual(kwargs["headers"]["x-goog-api-key"], "gemini-private-key")
        self.assertEqual(kwargs["json"]["model"], "gemini-test-model")
        self.assertEqual(kwargs["json"]["generation_config"]["max_output_tokens"], 4096)
        self.assertEqual(kwargs["timeout"], (10, 90))

    def test_gemini_interactions_multimodal_conversion(self):
        messages = [{"role": "user", "content": [
            {"type": "input_text", "text": "Planı incele"},
            {"type": "input_image", "image_url": "data:image/jpeg;base64,aGVsbG8="},
        ]}]
        converted = ag._to_gemini_interactions(ag._input_messages(messages))
        self.assertEqual(converted[0], {"type": "text", "text": "Planı incele"})
        self.assertEqual(converted[1], {"type": "image", "data": "aGVsbG8=", "mime_type": "image/jpeg"})

    def test_missing_provider_configuration_fails_closed(self):
        with self.assertRaises(ag.AIRouterError):
            ag.AIRouterClient({"AI_PROVIDER": "auto"})


if __name__ == "__main__":
    unittest.main()
