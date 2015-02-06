Imports FluentAssertions
Imports Rhino.Mocks


<TestClass()>
Public Class ComparaisonTest

    <TestMethod>
    Public Sub Guichet_Devrait_Interroger_Fonds_CompteBancaire()
        Const petitMontant As Integer = 125
        Dim guichet As Guichet

        ' Rhino Mocks
        Dim compteBancaireRhino = MockRepository.GenerateStub(Of ICompteBancaire)()
        compteBancaireRhino.Stub(Sub(x) x.FondsDisponibleSuffisant(petitMontant)).Return(True)

        ' NSubstitute
        Dim compteBancaireNSub = Substitute.For(Of ICompteBancaire)()
        compteBancaireNSub.FondsDisponibleSuffisant(petitMontant).Returns(True)

        ' Action et Vérification
        guichet = New Guichet(compteBancaireNSub)
        guichet.PeutRetirer(petitMontant).Should().BeTrue()
    End Sub

    <TestMethod>
    Public Sub Guichet_Devrait_Transmettre_Transaction_CompteBancaire()
        Const petitMontant As Integer = 125
        Dim guichet As Guichet

        ' Rhino Mocks
        Dim compteBancaireRhino = MockRepository.GenerateStub(Of ICompteBancaire)()
        compteBancaireRhino.Stub(Sub(x) x.DebiterCompte(petitMontant)).Return(True)

        ' NSubstitute
        Dim compteBancaireNSub = Substitute.For(Of ICompteBancaire)()
        compteBancaireNSub.DebiterCompte(petitMontant).Returns(True)

        ' Action
        guichet = New Guichet(compteBancaireNSub)
        guichet.Retirer(petitMontant).Should().BeTrue()

        ' Vérification NSubstitute
        compteBancaireNSub.Received(1).DebiterCompte(petitMontant)

        ' Vérification Rhino Mocks
        compteBancaireRhino.AssertWasCalled(Sub(x) x.DebiterCompte(petitMontant), Sub(x) x.Repeat.Once())
    End Sub

    <TestMethod()>
    Public Sub Substitution_Recursive()
        Dim compteBancaireRhino = MockRepository.GenerateStub(Of ICompteBancaire)()
        Dim compteBancaireNSub = Substitute.For(Of ICompteBancaire)()

        compteBancaireNSub.Client.Should().NotBeNull()
        compteBancaireRhino.Client.Should().NotBeNull("RhinoMock ne moque pas recursif")
    End Sub

    <TestMethod()>
    Public Sub Intercepter_Arguments_Complexe()
        Const petitMontant As Integer = 125
        Dim guichet As Guichet
        Dim client = Substitute.For(Of IClient)()
        Dim clientRecu As IClient

        ' Rhino Mocks
        Dim compteBancaireRhino = MockRepository.GenerateStub(Of ICompteBancaire)()
        compteBancaireRhino.Stub(Sub(x) x.Transferer(client, petitMontant)).WhenCalled(Sub(x) clientRecu = CType(x.Arguments.First(), IClient))

        ' NSubstitute
        Dim compteBancaireNSub = Substitute.For(Of ICompteBancaire)()
        compteBancaireNSub.Transferer(Arg.Do(Of IClient)(Sub(x) clientRecu = x), petitMontant)

        ' Action
        guichet = New Guichet(compteBancaireNSub)
        guichet.TransfererVersAutreClient(client, petitMontant)

        ' Vérification
        clientRecu.Nom.Should().BeSameAs(client.Nom)
    End Sub

    <TestMethod()>
    Public Sub Exemples_Divers()
        Dim autreClient = Substitute.For(Of IClient)()
        Dim compteBancaire = Substitute.For(Of ICompteBancaire)()

        ' Ignorer les arguments
        compteBancaire.FondsDisponibleSuffisant(0).ReturnsForAnyArgs(True)
        compteBancaire.FondsDisponibleSuffisant(1000).Should().BeTrue()

        ' Propriétés automatique et récursive
        compteBancaire.Client.Prenom = "Champion"
        compteBancaire.Client.Prenom.Should().Be("Champion")
        Dim client = compteBancaire.Client

        ' Contrainte des arguments
        compteBancaire.Transferer(Arg.Is(Of IClient)(Function(x) x.Prenom = client.Prenom),
                                  Arg.Any(Of Integer)).
                       Returns(True)
        compteBancaire.Transferer(autreClient, 100).Should().BeFalse("Contraintes ne correspondent pas")
        compteBancaire.Transferer(client, 10000000).Should().BeTrue("Contraintes correspondent")

        ' Appel non recue
        compteBancaire.DidNotReceive().Transferer(client, 0)
        compteBancaire.DidNotReceiveWithAnyArgs().DebiterCompte(0)
    End Sub
End Class

Public Class Guichet
    Private ReadOnly _compteBancaire As ICompteBancaire

    Public Sub New(ByVal compteBancaire As ICompteBancaire)
        _compteBancaire = compteBancaire
    End Sub

    Public Function PeutRetirer(ByVal montant As Integer) As Boolean
        Return _compteBancaire.FondsDisponibleSuffisant(montant)
    End Function

    Public Function Retirer(ByVal montant As Integer) As Boolean
        Return _compteBancaire.DebiterCompte(montant)
    End Function

    Public Sub TransfererVersAutreClient(ByVal client As IClient, ByVal montant As Integer)
        _compteBancaire.Transferer(client, montant)
    End Sub
End Class

Public Interface ICompteBancaire
    Function FondsDisponibleSuffisant(ByVal montant As Integer) As Boolean
    Function DebiterCompte(ByVal montant As Integer) As Boolean
    ReadOnly Property Client As IClient
    Function Transferer(ByVal autreClient As IClient, ByVal montant As Integer) As Boolean
End Interface

Public Interface IClient
    Property Nom() As String
    Property Prenom() As String
End Interface