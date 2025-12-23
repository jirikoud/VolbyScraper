namespace Volby.Model;

internal class MunicipalityModel
{
    public required string Name { get; set; }

    public required string County { get; set; }

    public required string Url { get; set; }

    public int UsedEnvelopes { get; set; }

    public int ValidVotes { get; set; }

    public required List<PartyModel> Parties { get; set; }
}

