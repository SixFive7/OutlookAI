using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookAI.Core.Com;

namespace OutlookAI.ComHost.Protocol
{
    /// <summary>
    /// Wire encoding for <see cref="ComDraftBody"/>.
    /// <para>
    /// Written by hand for two reasons. The type has a private constructor and only
    /// static factories, so the serializer cannot build it; and it exposes a computed
    /// <c>FormatName</c> that would otherwise round-trip as a phantom property.
    /// </para>
    /// <para>
    /// The alternative - annotating the type in Core - was rejected deliberately: Core
    /// also targets net48, where System.Text.Json is unavailable, and it carries an
    /// explicit no-JSON-dependency rule. Keeping the encoding here preserves that.
    /// </para>
    /// </summary>
    internal sealed class ComDraftBodyConverter : JsonConverter<ComDraftBody>
    {
        public override ComDraftBody? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.Null)
            {
                return null;
            }

            bool isHtml = false;
            string text = string.Empty;
            string html = string.Empty;

            if (reader.TokenType != JsonTokenType.StartObject)
            {
                throw new JsonException("Expected an object for ComDraftBody.");
            }

            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndObject)
                {
                    break;
                }

                if (reader.TokenType != JsonTokenType.PropertyName)
                {
                    continue;
                }

                string? name = reader.GetString();
                if (!reader.Read())
                {
                    break;
                }

                switch (name)
                {
                    case "isHtml":
                    case "IsHtml":
                        isHtml = reader.TokenType == JsonTokenType.True;
                        break;
                    case "text":
                    case "Text":
                        text = reader.GetString() ?? string.Empty;
                        break;
                    case "html":
                    case "Html":
                        html = reader.GetString() ?? string.Empty;
                        break;
                    default:
                        reader.Skip();
                        break;
                }
            }

            // Rebuild through the factories so the type's own invariant - exactly one of
            // Text/Html populated - is restored by its own code rather than asserted here.
            return isHtml ? ComDraftBody.FromHtml(html) : ComDraftBody.FromText(text);
        }

        public override void Write(Utf8JsonWriter writer, ComDraftBody value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            writer.WriteBoolean("isHtml", value.IsHtml);
            writer.WriteString("text", value.Text);
            writer.WriteString("html", value.Html);
            writer.WriteEndObject();
        }
    }
}
