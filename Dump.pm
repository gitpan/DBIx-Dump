package DBIx::Dump;

use 5.006;
use strict;
use warnings;

require Exporter;
#use AutoLoader qw(AUTOLOAD);

our @ISA = qw(Exporter);

# Items to export into callers namespace by default. Note: do not export
# names by default without a very good reason. Use EXPORT_OK instead.
# Do not simply export all your public functions/methods/constants.

# This allows declaration	use DBIx::Dump ':all';
# If you do not need this, moving things directly into @EXPORT or @EXPORT_OK
# will save memory.
our %EXPORT_TAGS = ( 'all' => [ qw(

) ] );

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } );

our @EXPORT = qw(

);
our $VERSION = '0.02';

sub new
{
	my $self = shift;
	my $attr = {@_};

	bless $attr, $self;
}

### Must put all anonymous subs before the %formats hash and dump sub.

my $excel = sub {

	my $self = shift;

	require Spreadsheet::WriteExcel;

	my $workbook = Spreadsheet::WriteExcel->new($self->{output});

	my $worksheet = $workbook->addworksheet();

	my $format = $workbook->addformat(); # Add a format
	$format->set_bold();
	$format->set_color('red');
	$format->set_align('center');

	my $col = 0; my $row = 0;
	my $cols = $self->{sth}->{NAME_uc};

	foreach my $data (@$cols)
	{
		$worksheet->write(0, $col, $data, $format);
		$col++;
	}
	$row++;
	$col = 0;

	while (my @data = $self->{sth}->fetchrow_array())
	{
		foreach my $data (@data)
		{
			$worksheet->write($row, $col, $data);
			$col++;
		}
		$col = 0;
		$row++;
	}
	$row = 0;
};

my $csv = sub {

	my $self = shift;

	require Text::CSV_XS;
	require IO::File;

	my $fh = IO::File->new("$self->{output}", "w");

	my $csvobj = Text::CSV_XS->new({
    'quote_char'  => '"',
    'escape_char' => '"',
    'sep_char'    => ',',
    'binary'      => 0
	});


	my $cols = $self->{sth}->{NAME_uc};
	$csvobj->combine(@$cols);
	print $fh $csvobj->string(), "\n";

	while (my @data = $self->{sth}->fetchrow_array())
	{
		$csvobj->combine(@data);
		print $fh $csvobj->string(), "\n";
	}
	$fh->close();
};

my %formats = (
								'excel' => $excel,
								'csv'		=> $csv
							);

sub dump
{
	my $self = shift;
	my $attr = {@_};
	$self = {%$self, %$attr};

	$formats{$self->{'format'}}->($self);
}


# Preloaded methods go here.

# Autoload methods go after =cut, and are processed by the autosplit program.

1;
__END__
# Below is stub documentation for your module. You better edit it!

=head1 NAME

DBIx::Dump - Perl extension for dumping database (DBI) data into a variety of formats.

=head1 SYNOPSIS

  use DBI;
	use DBIx::Dump;

	my $dbh = DBI->connect("dbi:Oracle:DSN_NAME", "user", "pass", {PrintError => 0, RaiseError => 1});
	my $sth = $dbh->prepare("select * from foo");
	$sth->execute();

	my $exceldb = DBIx::Dump->new('format' => 'excel', 'ouput' => 'db.xls', 'sth' => $sth);
	$exceldb->dump();

=head1 DESCRIPTION

DBIx::Dump allows you to easily dump database data, retrieved using DBI, into a variety of formats
including Excel, CSV, etc...

=head2 EXPORT

None by default.


=head1 AUTHOR

Ilya Sterin<lt>isterin@cpan.org<gt>

=head1 SEE ALSO

L<perl>.
L<DBI>.

=cut
