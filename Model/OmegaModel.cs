namespace Volby.Model;

internal class OmegaModel
{
    public required string Municipality {  get; set; }

    public required string County { get; set; }

    public required string Url { get; set; }

    public required int UsedEnvelopes { get; set; }

    public required int ValidVotes { get; set; }

    public required int OmegaVotes { get; set; }
}
