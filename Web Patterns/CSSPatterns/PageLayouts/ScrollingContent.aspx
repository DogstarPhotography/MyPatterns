<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ScrollingContent.aspx.vb"
	Inherits="CSSPatterns.ScrollingContent" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<title>Scrolling Content Box</title>
	<style type="text/css">
	<!--
	.ScrollBox 
	{
		height: 300px;
		width: 400px;
		padding: 5px;
		background-color: #CCCCFF;
		overflow: auto;
		}

	-->
	</style>
	<script src="ScrollingContent.js" type="text/javascript"></script>
</head>
<body>
<div id="ScrollNavigation">
<h2>Menu</h2>
<ul>
	<li><a href="#" onclick="ScrollToElementInDiv('LoremIpsumA');">Lorem Ipsum A</a></li>
	<li><a href="#" onclick="ScrollToElementInDiv('LoremIpsumB');">Lorem Ipsum B</a></li>
	<li><a href="#" onclick="ScrollToElementInDiv('LoremIpsumC');">Lorem Ipsum C</a></li>
	<li><a href="#" onclick="ScrollToElementInDiv('LoremIpsumD');">Lorem Ipsum D</a></li>
</ul>
</div>
	<div id="ScrollBoxDiv" class="ScrollBox">
		<p id="LoremIpsumA">
			Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur mollis dui eu
			libero bibendum vitae gravida nibh hendrerit. Mauris lobortis eros nec eros dignissim
			laoreet. Integer pharetra dictum lectus, sed convallis leo aliquet quis. Ut ac est
			lacus. Etiam adipiscing tortor eget diam lobortis mollis. Suspendisse placerat ligula
			aliquam tellus sollicitudin hendrerit. Pellentesque pharetra nisl id mi vulputate
			lacinia.</p>
		<p id="LoremIpsumB">
			Morbi nec tellus faucibus erat iaculis feugiat. Nulla sollicitudin volutpat turpis
			ut dapibus. Proin lectus dui, volutpat sed suscipit ut, adipiscing ac diam. Maecenas
			a elit augue. Aliquam ac lacus nec risus commodo hendrerit. Donec in orci ut nunc
			congue vehicula. Mauris vulputate erat id arcu tincidunt id malesuada neque lacinia.
			Fusce suscipit orci nec elit ornare dictum. Quisque arcu lorem, molestie vitae tincidunt
			id, imperdiet in mauris. Duis cursus lobortis leo, quis varius orci scelerisque
			eget. Phasellus eu pharetra arcu.</p>
		<p id="LoremIpsumC">
			Integer scelerisque, felis fringilla iaculis imperdiet, erat tellus pretium ante,
			id malesuada purus lorem ut mauris. Cras consectetur varius elit, vitae varius lectus
			lacinia tempus. Vestibulum tempor, felis in iaculis porttitor, lorem magna posuere
			nunc, sed tempus dolor elit eget diam. Proin gravida, nisi a bibendum pharetra,
			ante dui sagittis libero, a lacinia urna sapien non risus. Sed enim est, rhoncus
			at scelerisque ut, sollicitudin ut ligula. Maecenas euismod vehicula molestie. Pellentesque
			non ligula nec lectus vehicula lacinia vel euismod leo. Aliquam ultricies, dolor
			id semper interdum, ligula sapien aliquet augue, in vulputate lacus ante pulvinar
			diam. Praesent aliquet ultrices auctor. Quisque pulvinar arcu vitae mi viverra viverra.
			Donec sagittis, magna et scelerisque iaculis, ante quam placerat felis, ac imperdiet
			nulla libero non ipsum. Nulla non augue eu orci tristique egestas. Aliquam volutpat
			dolor ut mauris euismod placerat. Morbi volutpat lacus at metus rhoncus facilisis
			ac vitae massa.</p>
		<p id="LoremIpsumD">
			Ut magna ligula, scelerisque in fermentum a, tincidunt non dui. Sed erat libero,
			faucibus vel egestas nec, dignissim et massa. Donec egestas malesuada sem, ac consectetur
			neque hendrerit sed. Morbi eleifend, justo a commodo convallis, metus felis hendrerit
			felis, id condimentum diam arcu quis tortor. Curabitur faucibus bibendum mauris
			in tincidunt. Suspendisse blandit magna consectetur purus vestibulum placerat. Integer
			hendrerit dui sit amet lectus lobortis eleifend. Quisque eu sem at dolor cursus
			aliquam. Fusce eleifend dignissim faucibus. Curabitur tempus suscipit orci eget
			iaculis. Nulla facilisi. Nulla a purus est. Pellentesque habitant morbi tristique
			senectus et netus et malesuada fames ac turpis egestas. Vestibulum eu urna ante,
			in congue sapien. Donec gravida magna id tortor pellentesque ut pulvinar mi pellentesque.
			Pellentesque felis arcu, luctus eget dapibus vitae, ultrices eget arcu.</p>
	</div>
	<div id="BlockContent">
		<h2>Bunch of Lorem Ipsum to make the page longer</h2>
		<p>
			Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur mollis dui eu
			libero bibendum vitae gravida nibh hendrerit. Mauris lobortis eros nec eros dignissim
			laoreet. Integer pharetra dictum lectus, sed convallis leo aliquet quis. Ut ac est
			lacus. Etiam adipiscing tortor eget diam lobortis mollis. Suspendisse placerat ligula
			aliquam tellus sollicitudin hendrerit. Pellentesque pharetra nisl id mi vulputate
			lacinia.</p>
		<p>
			Morbi nec tellus faucibus erat iaculis feugiat. Nulla sollicitudin volutpat turpis
			ut dapibus. Proin lectus dui, volutpat sed suscipit ut, adipiscing ac diam. Maecenas
			a elit augue. Aliquam ac lacus nec risus commodo hendrerit. Donec in orci ut nunc
			congue vehicula. Mauris vulputate erat id arcu tincidunt id malesuada neque lacinia.
			Fusce suscipit orci nec elit ornare dictum. Quisque arcu lorem, molestie vitae tincidunt
			id, imperdiet in mauris. Duis cursus lobortis leo, quis varius orci scelerisque
			eget. Phasellus eu pharetra arcu.</p>
		<p>
			Integer scelerisque, felis fringilla iaculis imperdiet, erat tellus pretium ante,
			id malesuada purus lorem ut mauris. Cras consectetur varius elit, vitae varius lectus
			lacinia tempus. Vestibulum tempor, felis in iaculis porttitor, lorem magna posuere
			nunc, sed tempus dolor elit eget diam. Proin gravida, nisi a bibendum pharetra,
			ante dui sagittis libero, a lacinia urna sapien non risus. Sed enim est, rhoncus
			at scelerisque ut, sollicitudin ut ligula. Maecenas euismod vehicula molestie. Pellentesque
			non ligula nec lectus vehicula lacinia vel euismod leo. Aliquam ultricies, dolor
			id semper interdum, ligula sapien aliquet augue, in vulputate lacus ante pulvinar
			diam. Praesent aliquet ultrices auctor. Quisque pulvinar arcu vitae mi viverra viverra.
			Donec sagittis, magna et scelerisque iaculis, ante quam placerat felis, ac imperdiet
			nulla libero non ipsum. Nulla non augue eu orci tristique egestas. Aliquam volutpat
			dolor ut mauris euismod placerat. Morbi volutpat lacus at metus rhoncus facilisis
			ac vitae massa.</p>
		<p>
			Ut magna ligula, scelerisque in fermentum a, tincidunt non dui. Sed erat libero,
			faucibus vel egestas nec, dignissim et massa. Donec egestas malesuada sem, ac consectetur
			neque hendrerit sed. Morbi eleifend, justo a commodo convallis, metus felis hendrerit
			felis, id condimentum diam arcu quis tortor. Curabitur faucibus bibendum mauris
			in tincidunt. Suspendisse blandit magna consectetur purus vestibulum placerat. Integer
			hendrerit dui sit amet lectus lobortis eleifend. Quisque eu sem at dolor cursus
			aliquam. Fusce eleifend dignissim faucibus. Curabitur tempus suscipit orci eget
			iaculis. Nulla facilisi. Nulla a purus est. Pellentesque habitant morbi tristique
			senectus et netus et malesuada fames ac turpis egestas. Vestibulum eu urna ante,
			in congue sapien. Donec gravida magna id tortor pellentesque ut pulvinar mi pellentesque.
			Pellentesque felis arcu, luctus eget dapibus vitae, ultrices eget arcu.</p>
		<p>
			Proin congue faucibus commodo. Nulla ligula felis, porta eget tincidunt ac, semper
			at arcu. Proin egestas tortor eu urna vestibulum cursus. Curabitur non tellus velit,
			non suscipit nunc. Vivamus vitae nisl at neque elementum pellentesque. Etiam non
			felis enim, sed dignissim metus. Aliquam erat volutpat. Vestibulum lacinia massa
			sit amet orci malesuada rutrum. Sed malesuada erat non erat porta lobortis. Pellentesque
			habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.
			Cras vitae vulputate lacus. Sed augue velit, ultrices sit amet facilisis vitae,
			facilisis a risus. Aliquam nec risus nibh, eget fringilla lorem. Maecenas tincidunt
			ante tempus elit pretium tincidunt.</p>
	</div>
</body>
</html>
