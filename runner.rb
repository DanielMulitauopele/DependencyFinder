require './dependency_finder'

df = DependencyFinder.new

df.write_to_excel("opportunity_field_names.txt", [0, 1, 4, 12])